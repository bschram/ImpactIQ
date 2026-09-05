# Validation report - what the v2 audit found and what v3 does about it

Before v3 was written, the v2 monolith (`Legacy/Final PS Script (v2 interactive).txt`, 3 740 lines) and the four
C# extract scripts under `Config\` were read section by section by independent reviewers, and every high-severity
finding was re-verified by a second reviewer against the cited lines (Windows PowerShell 5.1 semantics, real-run
impact). This document lists those findings, where v3 addresses each one, which findings were refuted, and which
parts of v3 could only be validated with mocks on Linux and still need a run on a Windows box with real Power BI
access.

Finding ids: `C1..C9` = sections of the PowerShell script (bootstrap, auth, dialogs, scope, inventory, My Workspace +
model backup, report backup, report/model detail, dataflows); `X1..X4` = the C# scripts (`X1` Model Detail + Measure
Dependency, `X2`/`X3` classic `Report Detail Extract Script.csx`, `X4` `Report Detail Extract Script-PBIR.csx`).

## 1. Confirmed findings (verified) and where v3 fixes them

| Id | Severity | Finding (v2) | Fixed in v3 by |
|---|---|---|---|
| C1-04 | high | `$ErrorActionPreference = 'SilentlyContinue'` for errors AND warnings hid every bootstrap failure; "Downloaded" printed unconditionally | `ImpactIQ.ps1` sets `$ErrorActionPreference = 'Stop'`; every step is wrapped in try/catch with `Write-IQLog` (`ImpactIQ.Common.ps1`) which writes the console line and `Logs\ImpactIQ_*.log`; tool downloads log "Downloaded" only after success (`ImpactIQ.Tools.ps1` `Initialize-IQTools`); distinct exit codes 0 / 2 / 3 / 1 (`Get-IQExitCode`) |
| C1-10 | high | TE2 preflight never aborted, no timeout, no captured output, "files missing" confused with "exe failed" | `ImpactIQ.Tools.ps1`: `Invoke-IQProcess` (timeout, captured stdout/stderr under `tool-logs\`, process-tree kill) and the preflight in `Initialize-IQTools` with reasons `ok / missing-exe / missing-bim / timeout / exit N`; WER `DontShowUI=1` set in HKCU on Windows; `Invoke-IQTabularEditor` reused by every TE2 call (Models, Reports) |
| C4-03 | high | Workspace listing had no error handling; a failure looked like "Workspace selection cancelled" and looped forever | `ImpactIQ.Inventory.ps1` `Get-IQWorkspaceList` throws when the listing fails (paged `$top/$skip`); `Resolve-IQScope` / `ImpactIQ.Interactive.ps1` only treat the picker's own cancel message as "cancel", every other exception propagates |
| C5-01 | high | 11 unguarded `Invoke-PowerBIRestMethod` calls silently dropped whole sheets/workspaces | Every call goes through `Invoke-IQApi` (`ImpactIQ.Http.ps1`) with the retry matrix; `ImpactIQ.Inventory.ps1` `Invoke-IQInventoryGet` records failures per collector in the `ws-*.json` `Errors[]` (sheet `InventoryErrors`), core-collector failures mark the workspace item `Failed` so it is re-collected on resume |
| C5-02 | high | No 429 / throttling handling for thousands of sequential calls | `Invoke-IQApi`: 429 -> `Retry-After` (default 30 s, max 300 s, up to 8 retries); 5xx/408/network -> exponential back-off 2..60 s (`-MaxRetries`); 401 -> one forced token refresh and retry; Fabric paging honours `continuationToken`/`continuationUri`; call volume reduced per C5-09/C5-10 (Reports/Models mode filtering before per-item calls) |
| C6-01 | high | `Add-Member` passed `-NotePropertyValue` twice: My Workspace page rows lost `ReportId`/`ReportName` | `ImpactIQ.Inventory.ps1` `Get-IQMyWorkspaceInventory` adds `ReportId` and `ReportName` with `Add-IQNote` (verified by `Inventory.Tests.ps1`) |
| C6-03 | high | My Workspace REST calls had no error handling or retry | Same `Invoke-IQApi` wrapper + per-collector `Errors[]`; the `My Workspace` item is checkpointed like any workspace |
| C6-07 | high | TE2 exit code, output and per-dataset status discarded; XMLA failures invisible | `ImpactIQ.Models.ps1` `Invoke-IQModelBackupStage`: jobs run through `Invoke-IQProcessBatch`, `Complete-IQModelBackupJob` checks exit code, timeout and a non-empty `.bim`, records `Set-IQItemDone -Status Failed` with the first tool-log lines (manifest `failures`, `Failures` sheet), summary line "N exported (XMLA/Fabric), skipped, failed" |
| C7-01 | high | `$report.WebUrl` never existed: paginated reports mis-typed, RDL branch dead | `ImpactIQ.Reports.ps1` `Get-IQReportWorkList`: `IsPaginated = ReportType -eq 'PaginatedReport' -or ReportWebUrl -like '*/rdlreports/*'`; `.rdl` export path reachable; getDefinition fallback gated on `-not IsPaginated` |
| C7-02 | high | Pseudo-workspace ids ("My Workspace", "Shared Reports (No Workspace Access)") spliced into API URLs | `Export-IQReportUsingApi` uses the group-less `reports/{id}/Export` route when the workspace id is not a GUID; shared-no-access reports are `Skipped`; no Fabric fallback for pseudo workspaces |
| C7-03 | high | Export failure warnings silenced; no log or manifest | `Write-IQLog` everywhere (with `##vso[task.logissue]` under Azure DevOps); per-report checkpoint + `ReportExports.txt` (`Write-IQReportExportSummary`) |
| C7-04 | high | `Export-ReportDefinitionAsPbix` returned `$true` without verifying the `.pbix` | `Export-IQReportDefinitionAsPbix` / `Export-IQReportUsingApi` download to `<file>.partial`, verify existence and size > 0, delete zero-byte files, clean staging folders in `finally` |
| C8-01 | high | TE2 exit codes never checked; failures produced blank workbooks | `Invoke-IQTabularEditor` (exit code, timeout, script-error detection in stdout); `Invoke-IQReportDetailStage` checks the TXT outputs, `Invoke-IQModelDetailTabularEditor` checks both CSVs; failures are checkpointed `Failed` and surface as exit code 2 |
| C8-03 | high | Model/measure loops driven by in-memory names; Pro-workspace `.bim` files skipped, missing files launched | `Invoke-IQModelDetailStage` is driven by the `.bim` found per dataset (`Get-IQModelBimPath`: run folder, ModelBackup checkpoint, ReportBackup checkpoint `data.BimPath`); Pro `.bim` files are named by dataset (`<CleanWs> ~ <CleanDataset>.bim`, `Invoke-IQReportModelExtract`) and reused across reports; nothing is launched without a `.bim` |
| C9-01 | high | Gen1 dataflow loop had no error handling; failed exports wrote stale/empty backups | `ImpactIQ.Dataflows.ps1` `Export-IQGen1Dataflow`: pseudo workspaces skipped, empty body -> `Failed`, no zero-byte backup left, UTF-8 without BOM, dataflow list taken from the inventory |
| C9-02 | high | Fabric `getDefinition` 202 long-running response ignored (Gen2 backups dropped) | `Invoke-IQFabricLro` (`ImpactIQ.Http.ps1`) polls `Location`/`x-ms-operation-id` until `Succeeded` then reads `/result`; used by `Export-IQFabricDataflow`, `Export-IQReportDefinitionAsPbix` and `Save-IQModelDefinitionFromFabric` |
| C9-03 | high | Query regex could not handle nested-bracket attribute records, dotted/non-ASCII identifiers or `shared` inside strings | `Split-IQMSharedQuery` tokenizer (line-anchored `shared` header, `#"..."` names un-doubled, balanced attribute records attached as `Attributes`), covered by `Dataflows.Tests.ps1` |
| C9-06 | high | Today's dataflow folder wiped at start; results held in memory until the final Export-Excel | Per-dataflow checkpoints (`Set-IQItemDone`), extracts written immediately to `extracts\dataflows\<id>.json`, no folder wipe on resume; `Assemble` rebuilds `Dataflow Detail.xlsx` from the extracts (`-Stages Assemble`) |
| X1-03 | high | No CSV escaping in `Measure Dependency Extract Script.csx` | `FormatField` lambda added (every field quoted, `"` doubled) - `Config\Measure Dependency Extract Script.csx`; the DAX and Bim fallbacks use the same quoting (`Write-IQCsvFile`) |
| X1-05 | high | TE2 exit code ignored and missing `.bim` files still launched TE2 twice per dataset | `Get-IQModelDetailPlan` only schedules Tabular Editor when a `.bim` exists; exit codes / timeouts from `Invoke-IQProcessBatch`; per-model status in the checkpoint (`method`, `message`, `data`) |
| X1-08 | high | Pro models extracted only when a report had the same name as its dataset; Pro `ModelID` a string | Pro `.bim` named by dataset (see C8-03); `ModelID` = dataset GUID for dedicated models and the base name for Pro (PBIT join rule) in the Bim and DAX paths; DAX `INFO.VIEW.*` / `INFO.*` fallback keyed by `DatasetId` (`Get-IQModelDetailViaDax`) |
| X2-B1 / X3-B1 | high | `ReportLevelMeasures` rows had 10 columns against the 12-column header (classic script) | `Config\Report Detail Extract Script.csx`: rows go through `ReportLevelMeasures.Add(new ReportLevelMeasures {...})` and the 12-column writer; `DataType` / `DataCategory` read from `modelExtensions` when present |
| X2-R1 / X3-R1 | high | Classic script held all output until the end; one unhandled exception lost every report | Per-report `try/catch` (error rows in `ExtractErrors.txt`, `continue`), rows flushed and builders cleared after every report |
| X4-01 | high | PBIR script held everything in memory, wrote 11 files once at the end | Headers written before the loop, every report appended (`iqFlush`) the moment it is parsed; per-report `try/catch` |
| X4-03 | high | 48 empty catch blocks; failures invisible headless | Per-report and unzip failures logged to `ExtractErrors.txt` (`ReportName, Script, Stage, Error, ReportDate`) which `Assemble` turns into an `ExtractErrors` sheet; `Invoke-IQReportDetailStage` checks exit codes and output files and marks the item `Failed` (exit code 2) |

Additional high-severity items from the same audit that are not in the verified table but were fixed because the
brief requires them: `X3-B2` (ten `Expresssion` JSON paths corrected in the classic script), `X2-H1` / `X3-H2` /
`X4-04` (`IMPACTIQ_BASE`, `IMPACTIQ_DATE_FOLDER`, `IMPACTIQ_REPORT_DATE` honoured by all four csx scripts, set by
`Invoke-IQReportDetailStage` and `Invoke-IQModelDetailTabularEditor`), `X2-D1` / `X3-DG1` (report identity: each
export carries a `<name>.meta.json` sidecar and `ReportExports.txt`).

## 2. Medium / low findings honoured (unverified list, applied where the brief agrees)

| Area | Ids | v3 |
|---|---|---|
| Bootstrap / tools | C1-01, C1-03, C1-05, C1-06, C1-07, C1-08, C1-09, C1-17, C1-18, C1-13, C1-14 | `-BaseFolder` / `IMPACTIQ_BASE_FOLDER`; `Read-HostWithTimeout` guarded; `-SkipToolUpdate` / `IMPACTIQ_OFFLINE`; TLS 1.2 OR-ed in before downloads; progress bars off; `Config\tool-versions.json` stamp + `.partial` downloads; replace-instead-of-overlay extraction with rollback; system proxy with default credentials; `pbi-tools info` probe and Power BI Desktop detection; China `WebPrefix` typo; README timeout text |
| Auth | C2-01, C2-02, C2-04, C2-05, C2-12, C2-13 | Headless auth modes (`ImpactIQ.Auth.ps1`), no interactive re-login on token errors, JWT `exp`-based proactive refresh, `##vso[task.setsecret]` for minted tokens, environment table incl. Germany, UPN/tenant provenance in the manifest |
| Dialogs / scope | C3-02, C3-03, C3-06, C3-07, C4-01, C4-05, C4-07, C4-08, C4-10, C4-11 | Timers stopped after `ShowDialog`; "My Workspace only" scope; empty list is an error not a cancel; headless scope parameters with fail-fast "No scope"; inaccessible model workspaces warned and recorded; scan results cached in `$IQ.InventoryCache`; paged workspace listing |
| Inventory | C2-11, C5-04, C5-05, C5-06, C5-07, C5-09, C5-10, C5-12, C5-13, C5-14, C5-15, C5-16, C5-17 (parameters), C5-18, C2-14, C2-15 | Rename maps verbatim; report ids tracked across workspaces; `DataflowWorkspaceId` kept; Gen2 CI/CD dataflows via Fabric; `/transactions` skipped for CI/CD ids; mode filtering before per-item calls; nested values flattened to JSON; `RefreshHistoryTop` option; DirectQuery schedule + `DatasetRefreshScheduleKind`; new sheets Dashboards, DashboardTiles, WorkspaceUsers, DatasetUsers, DatasetParameters, Capacities; sensitivity label ids |
| Model backup / detail | C6-02, C6-06, C6-08, C6-10, C6-11, C6-12, C6-18, C6-19, C6-20, C8-02, C8-05, C8-06, C8-07, C8-10, C8-14, C8-16, X1-02, X1-06, X1-09, X1-10, X1-13, X1-19, X1-20, X1-21 | Pseudo workspaces never hit XMLA; no folder wipe on resume; per-job rename script; timeout + kill; clean-name collisions suffixed; `Model.Database.Name` set to the base name; per-model status in checkpoints; `Import-Csv -Encoding UTF8` (Assemble) and UTF-8 BOM in the fallback CSVs; `[ordered]` columns; skip-if-done; parallel pool; DAX/Bim fallback when TE2 cannot run; Fabric `getDefinition` (TMSL) fallback for the `.bim`; `Default` partition mode resolved; table `Description` filled |
| Report backup / detail | C7-06 .. C7-17, C8-04, C8-08, C8-09, X2-H2, X3-H1 | `subst` only for paths > 200 chars; `Invoke-IQWithRetry`/Http retries; pbi-tools + TE2 exit codes captured; name collisions suffixed; getDefinition archive not fed to pbi-tools; staging cleanup in `finally`; `ReportExports.txt`; optional `?format=PBIR`; header-only TXT handled; classic script only after the PBIR script; workbooks written to temp then moved |
| Dataflows | C9-04, C9-05, C9-07 .. C9-16 | Gen1 unescape only when still escaped; DataTable artefact columns written explicitly; Gen2 list from inventory; `.pq` byte-exact; `section` split fallback; 32 767-char cell guard (Assemble); `queriesMetadata`/`queryMetadata.json` joined; every Gen2 part saved; additive query columns; part-path validation |
| Assemble | C6-13, C6-14, C6-15, C6-16, C8-11, C8-15, C9-05, C9-11 | Single ExcelPackage per workbook written to a temp file then moved; union of columns; every contract sheet created even when empty (`Config\SheetContract.json`); `-AutoSize` only on Windows; streamed CSV merge; `-NoNumberConversion *` |

## 3. Refuted findings (not defects - no change made)

The verifier rejected these as findings against the v2 script because the described behaviour was either intended
(the monolith's "re-run today = start over" semantics) or a design limitation rather than a bug. v3 nevertheless
changed the design (checkpoints, no folder wipes on resume), so the situations they describe no longer exist:

* C3-08, C4-02, C5-03, C6-05 - run scope / inventory / extract collections held only in memory (v3 persists all of
  them under `State\runs\<RunId>\`).
* C6-06, C7-08, X1-09 - dated backup folders wiped on every run (v3 clears them only on a fresh run, never on resume).
* X1-04 - all-or-nothing csx extraction (v3 records per-model failures and falls back to the Bim parser / DAX).
* X2-R2, X3-R2, X4-02 - non-idempotent `File.Move` / append contract between the two report scripts (v3 removes the
  stale TXT files and unzip folders before every ReportDetail run, and runs the classic script only after the PBIR
  script succeeded).

## 4. What could NOT be validated on Linux - validate on Windows

The whole suite (`tests/Invoke-Tests.ps1`: parser, PSScriptAnalyzer with the 5.1 + 7.0 compatibility profiles, 300+
Pester tests with a mocked API) runs green on PowerShell 7.4 / Linux. The following paths are guarded with
`$script:IQ.IsWindows` / tool availability and were exercised only through fakes:

| Area | What is unverified | How to validate (Windows box, Windows PowerShell 5.1 and PowerShell 7) |
|---|---|---|
| Windows PowerShell 5.1 runtime | Everything ran on pwsh 7.4; 5.1-specific behaviour (`WebException` stream reads, `Start-Process` redirection, `Invoke-RestMethod` form bodies, `ConvertFrom-Json` limits) is coded from documentation | `powershell.exe -NoProfile -File tests\Invoke-Tests.ps1` (Pester 5.7.1 + ImportExcel installed); then a real run `powershell.exe -File ImpactIQ.ps1 -NonInteractive -Environment <env> -AuthMode DeviceCode -WorkspaceName '<small ws>'` |
| Tabular Editor 2 (XMLA export, csx scripts) | Real TE2 runs, XMLA connection string, csx behaviour incl. the edited scripts (`Expresssion` fix, `FormatField`, 12-column `ReportLevelMeasures`, per-report flush, `ExtractErrors.txt`, `IMPACTIQ_*` env vars) | Run `ModelBackup, ModelDetail` against one dedicated-capacity model and `ReportBackup, ReportDetail` against a few reports; compare `Model Detail.xlsx` / `Report Detail.xlsx` with a v2 run of the same objects; check `Logs\tool-logs\<Stage>\` for script compilation errors; open `ExtractErrors.txt` if present |
| pbi-tools (Pro `.bim` extraction) | Needs Power BI Desktop on the machine; the "no Desktop" detection is heuristic | Run `ReportBackup` on a Pro workspace with an IncludeModel export; expect `Model Backups\<RunId>\<Ws> ~ <Dataset>.bim`; without Desktop expect the Warn and the DAX/Bim fallback in `ModelDetail` |
| DPAPI token cache | `ProtectedData.Protect/Unprotect` (CurrentUser) - skipped test on Linux | DeviceCode sign-in twice under the same Windows user without `IMPACTIQ_TOKEN_CACHE_KEY`: the second run must not print a device code |
| WinForms dialogs (`ImpactIQ.Interactive.ps1`) | Moved verbatim; only their guards were exercised | Run `ImpactIQ.ps1` (or the launcher) interactively: environment dialog, run-mode dialog, pickers, cancel paths |
| Real REST / Fabric APIs | All responses were fixtures; paging, 429 and LRO behaviour are coded from Microsoft Learn; Fabric prefixes for GCC / GCC High / DoD are unverified (brief section 3) | First real run with `-Verbose` (`IMPACTIQ_DEBUG=1`): check the Debug HTTP lines, `InventoryErrors` and `Failures` sheets; for sovereign clouds confirm `FabricApiPrefix` (or set `-FabricApiPrefixOverride`) |
| DAX `INFO.*` over `executeQueries` | Microsoft documents raw `INFO.*` as unsupported there; `INFO.VIEW.*` is used first, raw parts best-effort | `-ModelDetailMethod Dax` on one model; inspect the checkpoint `data.UnavailableViaRest` / `DependencySource` and the CSVs |
| Fabric `getDefinition` for semantic models (TMSL) | Endpoint behaviour on Pro / shared capacity and GCC unknown | Watch the `ModelBackup` log for "Fabric getDefinition"; a circuit breaker disables it after 3 failures |
| Azure DevOps pipeline | YAML validated syntactically; task behaviour (artifact restore, `task.complete`, worktree push) not executed | Run `pipelines/azure-pipelines.yml` once with `stages: Inventory,Assemble` and `timeBudgetMinutes: 55` on a hosted agent; check the Summary tab, artifacts and exit-code mapping (0 / 2 / 3 / 1) |
| Time budget end-to-end | Unit-tested at the State/Common level (Paused stage, exit code 3, resume) | `ImpactIQ.ps1 ... -TimeBudgetMinutes 3` on a scope that needs longer: expect `Paused`, exit 3, partial workbooks; run again and expect completion with the finished items skipped |

## 5. How the validation was run here

* `pwsh -NoProfile -File tests/Invoke-Tests.ps1` (PowerShell 7.4.6, Linux): parser clean on every `.ps1`,
  PSScriptAnalyzer 0 errors and 0 `PSUseCompatibleSyntax` / `PSUseCompatibleCommands` findings (the remaining advisory
  warnings are `PSUseSingularNouns` on names the brief mandates and the WinForms `$sender` parameters), Pester green.
* Each module was additionally smoke-tested by its owner with strict mode (`Set-StrictMode -Version Latest`) and
  fake `TabularEditor.exe` / `pbi-tools.exe` shims to exercise the process pool, checkpoint and fallback paths.
* The C# scripts were edited textually (brace balance checked; no compiler available on Linux) - see section 4.
