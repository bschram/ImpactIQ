# ImpactIQ v3 - Automation overview

ImpactIQ v3 is the same tool as v2 (same four workbooks, same sheet and column names, same backup folder layout,
same `Power BI Governance Model.pbit`), re-built so that it can run **unattended** and **resume** after any
interruption. This page is the entry point; the other pages go deep on one topic each:

| Page | Read it when |
|---|---|
| [Auth-Options.md](Auth-Options.md) | you need to decide how a scheduled run signs in **without a service principal** (device code, cached refresh token, MFA-exempt account, Az context), or you are on a GCC / sovereign tenant |
| [Azure-DevOps.md](Azure-DevOps.md) | you want the included pipeline running: hosted vs self-hosted agent, secrets, artifacts, resume, and how the Power BI template reads the outputs from Azure Repos or SharePoint |
| [Headless-and-Resume.md](Headless-and-Resume.md) | you need the exact `ImpactIQ.ps1` parameters, environment variables, state layout, resume rules, exit codes and troubleshooting |
| [Data-Coverage.md](Data-Coverage.md) | you want to know which sheet comes from which API, what permission it needs, and what the optional `-IncludeAdminApis` / `-IncludeUsageMetrics` flags add |
| [Validation-Report.md](Validation-Report.md) | you want the audit findings behind the v3 changes |

Written for the reference deployment: **Power BI Premium (P SKU) on a GCC tenant, no Fabric capacity, no service
principal, no Azure VM, Azure DevOps available**. Everything also applies to commercial tenants (`-Environment Public`).

## 1. What "headless" and "resumable" mean here

* **Headless** - `ImpactIQ.ps1 -NonInteractive` never opens a dialog, never calls `Read-Host`, never opens a browser.
  Every choice the v2 pop-ups offered has a parameter (`-Environment`, `-RunMode`, `-WorkspaceName`, `-ReportId`,
  `-DatasetId`, `-IncludeMyWorkspace`, ...) or an environment variable (`IMPACTIQ_ENVIRONMENT`, `IMPACTIQ_USERNAME`,
  ...). When no scope is given headless, the run stops with "no scope" instead of scanning the whole tenant by accident.
* **Resumable** - the run is split into stages (`Inventory, ModelBackup, ReportBackup, ReportDetail, ModelDetail,
  Dataflows, Extras, Assemble`). Every workspace inventory, model backup, report download, model-detail CSV and dataflow
  export is checkpointed to `State\runs\<RunId>\done\...` the moment it finishes. Re-running the same command skips
  everything that already succeeded and retries only the failures. A run that dies at hour 5 of 6 loses minutes, not hours.
* **Robust** - one HTTP wrapper with retry/back-off (429 `Retry-After`, 5xx, network errors, one silent token refresh
  on 401), every external process (Tabular Editor 2, pbi-tools) has a timeout and captured output, every per-item
  failure is recorded in the manifest and the run continues. The exit code tells the scheduler what happened
  (`0` clean, `2` finished with item failures, `1` fatal).
* **Same outputs** - `Power BI Environment Detail.xlsx`, `Report Detail.xlsx`, `Model Detail.xlsx`,
  `Dataflow Detail.xlsx` in the base folder; backups under `Model Backups\<yyyy-MM-dd>\`, `Report Backups\<yyyy-MM-dd>\`,
  `Dataflow Backups\<yyyy-MM-dd>\`. New sheets and columns are additive only (see Data-Coverage.md).

The interactive experience is unchanged: double-click `Final PS Script` (now a thin launcher) or run `ImpactIQ.ps1`
without `-NonInteractive` and you get the environment prompt, the workspace/report/model pickers and the same folders.

## 2. Decision matrix - where to run it and how to sign in

No service principal means a **user identity** has to sign in. The constraints that decide everything:

* Microsoft-hosted agents get a **fresh VM per job** - nothing survives between runs except what you publish as an
  artifact. A device-code sign-in therefore only helps if the resulting refresh token is stored somewhere the next run
  can read: ImpactIQ encrypts it with your `IMPACTIQ_TOKEN_CACHE_KEY` and ships it in the `impactiq-state` artifact.
* Microsoft-hosted agents are capped at **60 min per job on the free tier, 360 min with one paid parallel job**.
  `-TimeBudgetMinutes` (pipeline parameter `timeBudgetMinutes`, e.g. `55`) makes ImpactIQ stop cleanly before the cap,
  build partial workbooks, exit `3` (`Paused`) and resume on the next run - so even the free tier works for big tenants,
  just over more runs.
  A large tenant takes longer - ImpactIQ simply resumes on the next run.
* Username/password (ROPC) sign-in **dies with MFA / Conditional Access / federation**, and Microsoft is deprecating
  it. It only works for a cloud-only, MFA-exempt account.
* Refresh tokens last **90 days of inactivity** (rolling) and are revoked by a password change, an admin
  "revoke sessions", or a Conditional Access policy that blocks device-code flow. Plan to re-bootstrap occasionally.

| # | Where | Auth mode | What you need | Pros | Cons | Fit for the reference tenant |
|---|---|---|---|---|---|---|
| A | Microsoft-hosted agent (`windows-latest`) | `Credential` (`IMPACTIQ_USERNAME` / `IMPACTIQ_PASSWORD`) | one MFA-/CA-exempt, cloud-only (PHS or PTA) Pro-licensed account; 1 paid parallel job for runs > 60 min | zero infrastructure, fully unattended, nothing to re-bootstrap | most GCC tenants enforce MFA/CA -> `AADSTS50076/50079/53003`; ROPC is deprecated; hosted IPs may be blocked by location policies; some agencies forbid hosted agents touching GCC data | only if security signs off on an MFA-exempt service account |
| B | Microsoft-hosted agent | `DeviceCode` + `IMPACTIQ_TOKEN_CACHE_KEY` (AES-encrypted refresh-token cache carried in the `impactiq-state` artifact) | 1 paid parallel job; one human completes the device-code sign-in on the first run (and again after ~90 idle days / password change / CA change); optional Teams/Slack webhook (`IMPACTIQ_DEVICECODE_WEBHOOK`) for the code | zero infrastructure, MFA-compatible (the human does MFA once), fully unattended afterwards | run retention deletes old artifacts (keep "days to keep runs" > the schedule gap); 360-min cap -> big tenants finish over 2-3 days; the cache key is a long-lived credential in Azure DevOps | **recommended when no Windows box is available** |
| C | Self-hosted agent on any existing Windows box (a workstation, a file server, a jump box) | `DeviceCode` with the default DPAPI cache **or** `AzContext` (`Connect-AzAccount -UseDeviceAuthentication` once, `Enable-AzContextAutosave`) | the Azure Pipelines agent installed as a service; the one-time sign-in done **as the agent's account** | no time cap, Tabular Editor / pbi-tools / DPAPI all native, cache never leaves the machine, no PAT or key to store in Azure DevOps, tenant-internal network | a machine that is on when the schedule fires; agent service account must be the same at bootstrap and run time; DPAPI cache is not portable | **recommended when any Windows box can host the agent** |
| D | Local Task Scheduler on a workstation/server (no Azure DevOps at all) | `DeviceCode` (DPAPI cache under the scheduled user) or `AzContext` | a scheduled task running `ImpactIQ.ps1 -NonInteractive ...` as a user who signed in once from the same account | simplest possible setup, no pipeline, logs and state stay local | no artifacts/history, outputs reach Power BI via a OneDrive-synced folder or a network share only; nobody is notified when it fails unless you add it | good for a single owner who already runs v2 by hand |

Quick reasoning for the reference tenant (GCC, P SKU, no Fabric):

1. **Try C first.** Any domain-joined Windows box that stays on can host the free self-hosted agent. It keeps every
   v2 behaviour (XMLA model backups through Tabular Editor 2, pbi-tools for Pro workspaces) and needs no secret in
   Azure DevOps at all.
2. **Otherwise B.** Buy one Microsoft-hosted parallel job (removes the 60-minute cap), set
   `IMPACTIQ_TOKEN_CACHE_KEY`, run the pipeline once by hand and complete the device-code sign-in from the run log or
   the webhook message. Subsequent scheduled runs are silent.
3. **A only with an MFA-exempt account** that your security team accepts, and expect it to stop working one day.
4. Fabric-only collectors (Fabric items, connections, gateways, Gen2 dataflows, `getDefinition` fallbacks) are
   **empty in GCC** because Fabric is not offered there; ImpactIQ degrades silently (one Debug line) and the Power BI
   REST collectors cover the rest. See Data-Coverage.md.

## 3. Five-minute setups

### C - self-hosted agent + DeviceCode (DPAPI cache)

```powershell
# 1. Install the agent as a service (Project settings > Agent pools > <pool> > New agent), e.g. under a domain service
#    account "svc-impactiq" that has a Pro license and Viewer/Contributor on the workspaces to document.
# 2. Bootstrap the token cache AS THAT ACCOUNT - the easiest way is a one-off manual run of the pipeline on that
#    agent (watch the log for the device code), or from an elevated prompt on the box:
runas /user:DOMAIN\svc-impactiq "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ImpactIQ\ImpactIQ.ps1 -BaseFolder C:\ImpactIQ -NonInteractive -Environment USGov -AuthMode DeviceCode -Stages Inventory -WorkspaceName Finance"
#    -> log shows "DEVICE CODE SIGN-IN REQUIRED: To sign in, use a web browser to open https://microsoft.com/devicelogin
#       and enter the code XXXXXXXXX" ; complete it within 15 minutes. State\auth\token-cache.json is now DPAPI-protected
#       for svc-impactiq on this machine.
# 3. Point pipelines/azure-pipelines.yml at the pool (pool parameter: { name: MyWindowsPool }), timeoutInMinutes: 0,
#    restoreState: false (the agent workspace persists), and schedule it.
```

### B - hosted agent + DeviceCode (AES cache key)

```powershell
# 1. Create a random 32+ character key and store it as SECRET pipeline variable IMPACTIQ_TOKEN_CACHE_KEY
[Convert]::ToBase64String((1..48 | ForEach-Object { Get-Random -Maximum 256 }) -as [byte[]])
# 2. Optional: store a Teams/Slack incoming-webhook URL as secret variable IMPACTIQ_DEVICECODE_WEBHOOK
# 3. Run the pipeline manually once (Run pipeline > authMode DeviceCode, environment USGov, workspaceNames "Finance*;HR").
#    Complete the device code shown in the "Run ImpactIQ" step log (or in the webhook message) within 15 minutes.
# 4. The run publishes State\auth\token-cache.json (AES-256, IV per file) inside impactiq-state; every later scheduled
#    run restores it and refreshes silently. Re-do step 3 if the log ever says "invalid_grant" / "new device-code flow".
```

### A - hosted agent + Credential

```text
Secret variables: IMPACTIQ_USERNAME = svc-impactiq@agency.gov, IMPACTIQ_PASSWORD = <password>
Run pipeline with authMode = Credential. If the log shows AADSTS50076/50079/53003/65001 the account is MFA/CA-bound:
switch to B or C (there is no workaround inside the tool).
```

### D - Task Scheduler

```powershell
# one-time, as the user that the task will run as:
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Power BI Backups\ImpactIQ.ps1" -BaseFolder "C:\Power BI Backups" -NonInteractive -Environment USGov -AuthMode DeviceCode -Stages Inventory -WorkspaceName "Finance"
# then the task action (daily 06:00, "run whether user is logged on or not", "do not store password" unchecked):
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Power BI Backups\ImpactIQ.ps1" -BaseFolder "C:\Power BI Backups" -NonInteractive -Environment USGov -AllWorkspaces -IncludeMyWorkspace
# exit code 0 = ok, 2 = finished with item failures (see Failures sheet), 3 = paused (-TimeBudgetMinutes reached; run it again), 1 = fatal (see Logs\ImpactIQ_*.log)
```

## 4. Environment variables (complete list)

Every variable is optional; parameters win over variables. Secrets are never logged and never written to the manifest.

| Variable | Used for | Equivalent parameter |
|---|---|---|
| `IMPACTIQ_ENVIRONMENT` | `Public`, `USGov`, `USGovHigh`, `USGovMil`, `China`, `Germany` (aliases `GCC`, `GCCHigh`, `DoD` accepted) | `-Environment` |
| `IMPACTIQ_USERNAME`, `IMPACTIQ_PASSWORD` | `Credential` mode (ROPC). Both set -> `Auto` picks `Credential` | `-Credential` |
| `IMPACTIQ_PBI_TOKEN` (+ `IMPACTIQ_FABRIC_TOKEN`) | `AccessToken` mode: a bearer token minted elsewhere; no refresh possible | `-AuthMode AccessToken` |
| `IMPACTIQ_TOKEN_CACHE_KEY` | AES-256 key for `State\auth\token-cache.json` (DeviceCode). Without it Windows uses DPAPI, other OS do not persist | `-TokenCacheKey` |
| `IMPACTIQ_TOKEN_CACHE_PATH` | alternative cache file location (default `<BaseFolder>\State\auth\token-cache.json`) | `-TokenCachePath` |
| `IMPACTIQ_DEVICECODE_WEBHOOK` | Teams/Slack incoming webhook that receives the device-code message as `{ "text": "..." }` | `-DeviceCodeWebhookUrl` |
| `IMPACTIQ_TENANT_ID` | tenant GUID or domain instead of `organizations` (guest accounts, multi-tenant users) | `-TenantId` |
| `IMPACTIQ_CLIENT_ID` | public client id used for device code / ROPC (default Azure PowerShell `1950a258-227b-4e31-a9cf-717495945fc2`; alternative Azure CLI `04b07795-8ddb-461a-bbee-02f9e1bf7b46`) | `-ClientId` |
| `IMPACTIQ_OFFLINE=1` | never download Tabular Editor 2 / pbi-tools updates (same as `-SkipToolUpdate`) | `-SkipToolUpdate` |
| `IMPACTIQ_DEBUG=1` | echo Debug-level log lines to the console (they are always in the log file) | `-Verbose` |
| `TF_BUILD`, `CI` | set by Azure DevOps / CI systems; force non-interactive mode and `##vso` log issues | `-NonInteractive` |

## 5. What to expect the first time

1. `Initialize-IQTools` checks the committed Tabular Editor 2 portable and pbi-tools under `Config\`, downloads
   updates unless `-SkipToolUpdate`, and runs a 2-minute preflight. On a non-Windows host (or when the preflight fails)
   model backups fall back to the DAX `INFO.*` path for model detail and report the reason per item.
2. `Initialize-IQAuth` resolves the mode (`Auto` order: Credential -> AccessToken -> cached DeviceCode -> Interactive ->
   DeviceCode) and signs in. In DeviceCode mode the message is logged at **Warn** level so it stands out in pipeline logs.
3. `Initialize-IQRun` decides between a fresh run and a resume (rules in Headless-and-Resume.md section 4) and writes
   `State\runs\<RunId>\manifest.json`.
4. Stages run in order, each item checkpointed. The console (and the log) ends with a per-stage summary table, the four
   output paths and the exit code.
5. Open `Power BI Governance Model.pbit`, point `Base Directory` at the folder (or set `UseWeb` for SharePoint /
   Azure Repos, see Azure-DevOps.md section 6) and refresh.
