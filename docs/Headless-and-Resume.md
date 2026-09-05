# Headless runs, state and resume

Everything `ImpactIQ.ps1` does unattended: parameters, environment variables, what is written where, how a re-run
decides what to skip, exit codes, logging and troubleshooting. Works on Windows PowerShell 5.1 and PowerShell 7.

Contents: 1 parameters - 2 examples - 3 state layout and manifest - 4 resume rules - 5 stages and checkpoints -
6 scope resolution - 7 exit codes - 8 logging - 9 troubleshooting.

## 1. `ImpactIQ.ps1` parameters

| Parameter | Default | Meaning |
|---|---|---|
| `-BaseFolder <path>` | the script folder when it contains `Config\`, else `C:\Power BI Backups` | root for `Config\`, `State\`, `Logs\`, the backup folders and the four workbooks |
| `-Environment Public\|Germany\|USGov\|China\|USGovHigh\|USGovMil` | `IMPACTIQ_ENVIRONMENT`, else the interactive dialog (60 s -> Public), else `Public` | cloud endpoints (Auth-Options.md section 3) |
| `-AuthMode Auto\|Interactive\|DeviceCode\|Credential\|AzContext\|AccessToken` | `Auto` | Auth-Options.md |
| `-TenantId <guid or domain>` | `organizations` | tenant for device code / ROPC |
| `-ClientId <guid>` | `1950a258-227b-4e31-a9cf-717495945fc2` | public client for device code / ROPC |
| `-Credential <PSCredential>` | `IMPACTIQ_USERNAME` / `IMPACTIQ_PASSWORD` | Credential mode |
| `-TokenCachePath <file>` | `<BaseFolder>\State\auth\token-cache.json` | DeviceCode refresh-token cache |
| `-TokenCacheKey <string>` | `IMPACTIQ_TOKEN_CACHE_KEY` | AES key for the cache (else DPAPI on Windows) |
| `-DeviceCodeWebhookUrl <url>` | `IMPACTIQ_DEVICECODE_WEBHOOK` | Teams/Slack webhook for the device-code message |
| `-AuthorityOverride <url>`, `-FabricApiPrefixOverride <url>` | table values | endpoint overrides |
| `-NonInteractive` | off (auto-on under `TF_BUILD` / `CI`) | no dialogs, no `Read-Host`, no browser; "no scope" becomes an error instead of "scan everything" |
| `-RunMode Workspaces\|Reports\|Models` | `Workspaces` | what the scope parameters mean (section 6) |
| `-WorkspaceId <guid[]>`, `-WorkspaceName <pattern[]>`, `-AllWorkspaces`, `-IncludeMyWorkspace` | none | workspace selection (`-WorkspaceName` uses `-like` wildcards) |
| `-ReportId <guid[]>`, `-DatasetId <guid[]>` | none | required in `Reports` / `Models` mode |
| `-Stages <name[]>`, `-SkipStages <name[]>` | all | `Inventory, ModelBackup, ReportBackup, ReportDetail, ModelDetail, Dataflows, Extras, Assemble` |
| `-RunId <string>` | today `yyyy-MM-dd` | run/backup folder name (`Model Backups\<RunId>\` ...). Keep the date format so the csx scripts and the PBIT's `ModelAsOfDate` keep working |
| `-Resume Auto\|Always\|Never` | `Auto` | section 4 |
| `-ResumeMaxAgeDays <int>` | `3` | how far back `Auto` looks for an unfinished run |
| `-Force` | off | delete `State\runs\<RunId>` and the three `<RunId>` backup folders, then start fresh |
| `-RefreshInventory` | off | on a resume, re-run the `Inventory` stage even though it completed |
| `-ModelDetailMethod Auto\|TabularEditor\|Dax\|Both` | `Auto` | `Auto` = Tabular Editor csx when a `.bim` exists and TE2 works, else DAX `INFO.*` over `executeQueries`; `Both` = TE2 then DAX on failure |
| `-MaxParallelExtracts <int>` | `2` | concurrent Tabular Editor / pbi-tools processes |
| `-ToolTimeoutMinutes <int>` | `20` | per external process (ReportDetail uses 3x) |
| `-MaxRetries <int>` | `5` | HTTP retries for 5xx/network errors (429 has its own 8-retry budget honouring `Retry-After`) |
| `-SkipToolUpdate` | off (`IMPACTIQ_OFFLINE=1`) | do not download Tabular Editor 2 / pbi-tools updates |
| `-IncludeAdminApis` | off | Extras: admin groups, Scanner API, activity events (Fabric admin only; probed, else skipped) |
| `-IncludeUsageMetrics` | off | Extras: per-workspace "Usage Metrics Report" model via DAX |
| `-ActivityDays <int>` | `30` | days of `admin/activityevents` (API window is 28) |
| `-LogPath <file>` | `Logs\ImpactIQ_yyyyMMdd_HHmmss.log` | log file |
| `-PassThru` | off | return the manifest object |

Interactive-only behaviours (never in headless mode): environment dialog, run-mode dialog, workspace/report/model
pickers, "My Workspace" included when the picker times out, and the pickers' 60-second timeouts.

## 2. Examples

```powershell
# Everything the account can see, GCC, unattended, resume if yesterday died
.\ImpactIQ.ps1 -BaseFolder 'C:\ImpactIQ' -NonInteractive -Environment USGov -AllWorkspaces -IncludeMyWorkspace

# Two workspace families, only inventory + dataflows, DAX-only model detail (no Tabular Editor needed)
.\ImpactIQ.ps1 -BaseFolder 'C:\ImpactIQ' -NonInteractive -Environment USGov -WorkspaceName 'Finance*','HR' `
    -Stages Inventory,Dataflows,ModelDetail,Assemble -ModelDetailMethod Dax

# Specific reports (their models and model workspaces are added automatically)
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -RunMode Reports -ReportId 6d5e...,0a1b...

# Specific models (their reports across all workspaces are added automatically)
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -RunMode Models -DatasetId 9f2c...

# Re-run only the failed model backups of run 2026-09-04 (checkpointed successes are skipped)
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -RunId 2026-09-04 -Resume Always -Stages ModelBackup,ModelDetail,Assemble

# Rebuild the four workbooks from the state of the latest run without any API call
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -Resume Always -Stages Assemble

# Start over today, wiping today's state and backup folders
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -AllWorkspaces -Force

# Extras (needs Fabric admin for the admin APIs; usage metrics need Contributor+ and the tenant setting)
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -AllWorkspaces -IncludeAdminApis -IncludeUsageMetrics -ActivityDays 7

# Windows PowerShell 5.1 from Task Scheduler / cmd
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\ImpactIQ\ImpactIQ.ps1" -BaseFolder "C:\ImpactIQ" -NonInteractive -Environment USGov -AllWorkspaces
```

Environment-variable equivalents (parameters win): `IMPACTIQ_ENVIRONMENT`, `IMPACTIQ_USERNAME`, `IMPACTIQ_PASSWORD`,
`IMPACTIQ_PBI_TOKEN`, `IMPACTIQ_FABRIC_TOKEN`, `IMPACTIQ_TOKEN_CACHE_KEY`, `IMPACTIQ_TOKEN_CACHE_PATH`,
`IMPACTIQ_DEVICECODE_WEBHOOK`, `IMPACTIQ_TENANT_ID`, `IMPACTIQ_CLIENT_ID`, `IMPACTIQ_OFFLINE`, `IMPACTIQ_DEBUG`.

## 3. State layout and manifest

```
<BaseFolder>\
  Config\                                  csx scripts, Blank Model.bim, TabularEditor\, PBI Tools\, Modules\, SheetContract.json, Temp\
  Logs\ImpactIQ_yyyyMMdd_HHmmss.log        one log per invocation (Debug lines always included)
  Logs\tool-logs\tools\                    Tabular Editor / pbi-tools preflight output
  State\auth\token-cache.json              encrypted refresh token (DeviceCode) - never inside a run, never an output
  State\runs\<RunId>\manifest.json         the run manifest (below); manifest.<timestamp>.json = archived completed run
  State\runs\<RunId>\inventory\workspaces.json     selected workspaces (incl. pseudo "My Workspace")
  State\runs\<RunId>\inventory\global.json         apps, app reports, connections, gateways, capacities
  State\runs\<RunId>\inventory\ws-<workspaceId>.json  everything collected for one workspace
  State\runs\<RunId>\inventory\extras-<sheet>.json    Extras sheets (admin / scanner / activity / usage)
  State\runs\<RunId>\done\<Stage>\<itemKey>.json      per-item checkpoint {status, outputs[], method, message, data}
  State\runs\<RunId>\extracts\dax\<datasetId>\*.json  raw INFO.* results (reused on re-run without re-querying)
  State\runs\<RunId>\extracts\dataflows\<id>.json     parsed dataflow queries (Dataflow Detail.xlsx is built from these)
  State\runs\<RunId>\extracts\admin\ , extracts\usage\  raw Extras responses
  State\runs\<RunId>\tool-logs\<Stage>\<itemKey>.out.txt / .err.txt  external process output
  Model Backups\<RunId>\<Ws> ~ <Model>.bim, .csv, _MD.csv
  Report Backups\<RunId>\<Ws> ~ <Report>.pbix|.rdl, *.txt (report detail extracts), ReportExports.txt
  Dataflow Backups\<RunId>\<Ws> ~ <Dataflow>.txt|.pq, <name>.definition\, Dataflow Detail.xlsx
  Power BI Environment Detail.xlsx, Report Detail.xlsx, Model Detail.xlsx, Dataflow Detail.xlsx
```

`manifest.json` (schema 1, no secrets):

```json
{ "schemaVersion": 1, "runId": "2026-09-04", "status": "Running|Completed|CompletedWithErrors|Failed",
  "startedUtc": "...", "updatedUtc": "...", "endedUtc": null,
  "host": { "machine": "", "user": "", "psVersion": "", "isAzureDevOps": false },
  "auth": "DeviceCode (cached) as user@agency.gov", "environment": "USGov",
  "options": { "...effective parameters, secrets removed..." },
  "scope": { "runMode": "Workspaces", "workspaceIds": [], "reportIds": [], "datasetIds": [], "includeMyWorkspace": false },
  "stages": { "Inventory": { "status": "Completed", "startedUtc": "", "endedUtc": "", "itemsDone": 12, "itemsFailed": 0, "error": null }, "...": {} },
  "failures": [ { "stage": "ModelBackup", "itemKey": "<datasetId>", "item": "WS ~ Model", "message": "...", "timeUtc": "..." } ],
  "outputs": { "environmentWorkbook": "...", "reportWorkbook": "...", "modelWorkbook": "...", "dataflowWorkbook": "..." },
  "resumes": [], "resumeCount": 0 }
```

The same data lands in the `RunSummary` and `Failures` sheets of `Power BI Environment Detail.xlsx`. Every JSON write is
atomic (`.tmp` + move), so a crash never leaves a half-written manifest or checkpoint.

## 4. Resume rules (`Initialize-IQRun`)

Effective `RunId` = `-RunId` if given, else today's `yyyy-MM-dd`.

| `-Resume` | Behaviour |
|---|---|
| `-Force` (any) | delete `State\runs\<RunId>` and clear `Model Backups\<RunId>`, `Report Backups\<RunId>`, `Dataflow Backups\<RunId>`; fresh run |
| `Never` | fresh run for `<RunId>` (same as `-Force` for that RunId; other runs' backups untouched) |
| `Always` | resume `<RunId>` if its manifest exists (any status, even `Completed`), else fresh |
| `Auto` (default) | 1. manifest for `<RunId>` exists and status != `Completed` -> resume it; 2. else, when `-RunId` was **not** given, the newest manifest with status `Running` / `Failed` / `CompletedWithErrors` whose `startedUtc` is within `-ResumeMaxAgeDays` (3) -> resume **that** RunId (yesterday's run that died is finished today, in yesterday's folders); 3. else fresh run for `<RunId>` - an existing `Completed` manifest for today is archived to `manifest.<timestamp>.json` and today's backup folders are cleared (the v2 "re-run today = start over" behaviour) |

On a resume:

* nothing is deleted; `status` goes back to `Running`, `resumeCount` increments, stages left `Running` by a crash are
  marked `Interrupted` and re-run;
* stages already `Completed` are skipped ("skipping" in the log) - except `Assemble`, which always rebuilds the four
  workbooks, and `Inventory` when `-RefreshInventory` is given (its checkpoints are deleted so every workspace is
  re-collected);
* inside a stage, items whose checkpoint is `Succeeded` or `Skipped` **and whose output files still exist** are skipped;
  a checkpoint whose `.bim`/`.pbix`/`.csv` was deleted is treated as not done;
* failed items are retried; when they succeed, their entries are removed from `manifest.failures`;
* the persisted `scope` is reused; differing scope parameters on the command line are ignored with a Warn (use `-Force`
  or a new `-RunId` to change scope);
* the DAX fallback rebuilds CSVs from `extracts\dax\` without re-querying; dataflow workbooks rebuild from
  `extracts\dataflows\`.

Running two instances against the same `BaseFolder` at the same time is not supported (checkpoints are per file, but the
manifest is shared).

## 5. Stages and checkpoints

| Stage | Item key | Skipped when | Output(s) in the checkpoint | Notes |
|---|---|---|---|---|
| `Inventory` (fatal) | workspace id (+ `global`, `My Workspace`) | `done\Inventory\<id>.json` and the `ws-*.json` exist | the `ws-<id>.json` file | a partial workspace (one collector threw) is `Failed` and re-collected on resume |
| `ModelBackup` | dataset id | `.bim` exists and is > 0 bytes | `Model Backups\<RunId>\<Ws> ~ <Model>.bim` | dedicated capacity only (XMLA via Tabular Editor 2, `-MaxParallelExtracts` in parallel); Pro workspaces are `Skipped` (model comes from the PBIX in ReportBackup); TE2 unavailable -> `Failed` with reason, ModelDetail still runs via DAX |
| `ReportBackup` | report id | file exists | `.pbix` / `.rdl` (+ `.bim` for Pro IncludeModel exports) | Export API (`IncludeModel` for Pro, `LiveConnect` for dedicated), Fabric `getDefinition` fallback where Fabric exists, pbi-tools extraction for Pro; `ReportExports.txt` summarises method per report |
| `ReportDetail` | `all` | the stage completed | the `*.txt` extract files | two csx scripts run once over every PBIX of the run folder (timeout 3x `-ToolTimeoutMinutes`) |
| `ModelDetail` | dataset id | both CSVs exist | `<Ws> ~ <Model>.csv`, `_MD.csv` | method recorded (`TabularEditor` / `Dax`) |
| `Dataflows` | dataflow id | backup + extract exist | `.txt` (Gen1) / `.pq` (Gen2) + `extracts\dataflows\<id>.json` | |
| `Extras` | `admin-groups`, `admin-scan`, `admin-activity-<date>`, `usage-<workspaceId>` | each item | `inventory\extras-*.json` | only with `-IncludeAdminApis` / `-IncludeUsageMetrics` |
| `Assemble` | the four workbooks | never (always rebuilt) | the workbook paths in `manifest.outputs` | writes to a temp file then moves, so a half-written workbook never replaces a good one; every sheet in `Config\SheetContract.json` exists even when empty |

`-Stages` / `-SkipStages` select stages; `Inventory` is required by every later stage on a fresh run (on a resume its
files are already there, so `-Stages ModelBackup` alone works).

## 6. Scope resolution (headless)

| `-RunMode` | Required | What is selected |
|---|---|---|
| `Workspaces` | `-WorkspaceId` and/or `-WorkspaceName` (wildcards) or `-AllWorkspaces`; `-IncludeMyWorkspace` optional | the matching workspaces (unknown ids/names -> Warn). Nothing selected and no `-IncludeMyWorkspace` -> **"No scope" error, exit 1** (never scans everything by accident) |
| `Reports` | `-ReportId` | the candidate workspaces (`-WorkspaceId/-WorkspaceName` if given, else all) are scanned for those reports; each report's dataset workspace is added; only the selected reports and their datasets are processed |
| `Models` | `-DatasetId` | the models' workspaces plus every workspace whose reports use those datasets; only those datasets and reports are processed |

The resolved scope is persisted in `manifest.scope` and reused on resume. Interactive runs use the v2 dialogs and
produce the same structure.

## 7. Exit codes

| Code | Meaning | Manifest status | Pipeline result |
|---|---|---|---|
| `0` | every stage `Completed`; or the user cancelled the interactive scope dialog | `Completed` | Succeeded |
| `2` | finished, but at least one item failed or a non-fatal stage failed (details in `manifest.failures` / the `Failures` sheet); outputs were still produced | `CompletedWithErrors` | SucceededWithIssues |
| `1` | fatal: authentication, no scope, `Inventory` stage failure, tool bootstrap when nothing can run, or an unhandled exception | `Failed` | Failed |

`-PassThru` returns the manifest object in addition to setting the exit code.

## 8. Logging

* Console and file: `[HH:mm:ss] [LEVEL] [Stage] [Item] message`; levels Info, Warn, Error, Debug, Success. Debug lines
  are always in the file, echoed to the console with `-Verbose` or `IMPACTIQ_DEBUG=1`.
* Under Azure DevOps (`TF_BUILD=True`) Warn/Error lines are also emitted as `##vso[task.logissue type=warning|error]`,
  so they show up in the run summary.
* Redaction: `Password=...` (XMLA connection strings), `Bearer ...`, `access_token=`, `refresh_token=`,
  `client_secret=`, `password=` are replaced by `***` in every line. Tokens minted on Azure DevOps are registered as
  secrets with the agent.
* Every HTTP call is logged at Debug as `METHOD path -> status (ms)`; `manifest.json` does not carry them, but
  `$IQ.Stats` (ApiCalls, Retries) is printed in the final summary.
* External processes: `tool-logs\<Stage>\<item>.out.txt/.err.txt`; the first error lines are copied into the item's
  checkpoint message and the `Failures` sheet.

## 9. Troubleshooting

| Symptom | Where to look | Fix |
|---|---|---|
| `No scope` / exit 1 immediately | first Error line | pass `-WorkspaceName`, `-WorkspaceId`, `-AllWorkspaces`, `-ReportId` or `-DatasetId` |
| `Interactive authentication is not possible in a non-interactive session` | Auth lines | set an auth mode that works headless: DeviceCode (+ cache), Credential, AzContext, AccessToken (Auth-Options.md) |
| `AADSTS50076/50079/53003` | Auth lines | MFA/CA: use DeviceCode |
| `DEVICE CODE SIGN-IN REQUIRED` every run | Auth lines: `Token cache ... not persisted` / `wrong key` | set `IMPACTIQ_TOKEN_CACHE_KEY` (non-Windows or hosted agents) or keep the same Windows user (DPAPI) |
| `429` with long waits | Debug HTTP lines | normal; the tool sleeps `Retry-After` (max 300 s, 8 retries). Lower `-MaxParallelExtracts`, split the scope across schedules |
| `ModelBackup` items `Failed: ... XMLA endpoint ... read-only` or `The XMLA endpoint is disabled` | Failures sheet / tool log | capacity setting *XMLA Endpoint* must be Read or Read Write and the tenant setting *Allow XMLA endpoints and Analyze in Excel* on; Build permission on the model. ModelDetail still works via DAX |
| `ModelBackup` `Failed: Tabular Editor preflight ...` | `Logs\tool-logs\tools\` | non-Windows host, missing `Config\TabularEditor\TabularEditor.exe`, blocked download (`-SkipToolUpdate`), or .NET Framework 4.7.2+ missing on the agent |
| `TimedOut` on Tabular Editor / pbi-tools | tool-logs | raise `-ToolTimeoutMinutes`, lower `-MaxParallelExtracts`; large models over XMLA can take > 20 min |
| `ModelDetail` `Failed: HTTP 400 ... INFO functions ... not supported` (`3239575574`) | Failures sheet | the tenant/capacity rejects raw `INFO.*` over `executeQueries` (Microsoft documents it as unsupported; it works on many tenants). Use Tabular Editor (`-ModelDetailMethod TabularEditor` on a dedicated capacity) or ask Microsoft support to enable INFO queries for the capacity |
| `ModelDetail` `Failed: HTTP 401/403` on `executeQueries` | Failures sheet | tenant setting *Semantic Model Execute Queries REST API* (Integration settings) must be on; the account needs **Build** on the model (Contributor+ has it). Measure expressions are blank without Write (Contributor+) |
| `ReportBackup` `Failed: 403` / `Export` errors | Failures sheet | tenant setting *Download reports* off, or the account is Viewer (needs Contributor+); large-storage-format / Direct Lake / incremental-refresh models cannot be exported with model - the tool falls back to `getDefinition` where Fabric exists |
| `Dataflows` Gen2 items `Failed: Fabric getDefinition returned nothing` | Failures sheet | no Fabric token (GCC), no read+write on the dataflow, or an encrypted sensitivity label |
| Workbook missing a sheet the PBIT expects | Assemble Warn lines | every contract sheet is created empty; if the PBIT still complains, run `-Stages Assemble -Resume Always` and check `Config\SheetContract.json` is present |
| Run resumed the wrong day | first lines: `Resuming run <RunId>` | pass `-RunId` explicitly, or `-Force` for a clean start |
| `Set-StrictMode` / `$IsWindows` errors on PowerShell 5.1 | log | report it - modules avoid PS7-only syntax and use `$script:IQ.IsWindows`; `Install-Module ImportExcel` is the only external dependency |
