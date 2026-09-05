# Running ImpactIQ from Azure DevOps Pipelines

The repository ships a ready pipeline: `pipelines/azure-pipelines.yml` (schedule + parameters) and the reusable steps
template `pipelines/templates/impactiq-run.yml`. This page explains how to wire it up for the reference tenant
(GCC, Premium P SKU, no service principal) and how the Power BI template reads the results afterwards.

Contents: 1 prerequisites - 2 create the pipeline - 3 secrets - 4 hosted vs self-hosted agents - 5 how a run flows
(state, artifacts, resume) - 6 getting the outputs into Power BI - 7 SharePoint / OneDrive - 8 pipeline parameters -
9 troubleshooting.

## 1. Prerequisites

* An Azure DevOps project with this repository in **Azure Repos** (or GitHub connected to Azure Pipelines - the
  "commit outputs" step assumes Azure Repos).
* A user account for the runs: Pro or PPU license, Viewer (metadata) / Contributor (backups) on the workspaces to
  document. Nothing else - no app registration.
* Parallel jobs: the **free Microsoft-hosted tier caps a job at 60 minutes**; buy one Microsoft-hosted parallel job to
  lift it to 360 minutes, or use the free self-hosted parallel job (no cap). Organization settings > Parallel jobs.
* The repository root is the ImpactIQ base folder (`ImpactIQ.ps1`, `Config\` with the csx scripts, `Blank Model.bim`,
  Tabular Editor 2 portable and pbi-tools). The pipeline passes `-BaseFolder $(Build.SourcesDirectory)`.

## 2. Create the pipeline

1. Pipelines > New pipeline > Azure Repos Git > this repo > **Existing Azure Pipelines YAML file** >
   `/pipelines/azure-pipelines.yml` > Save (do not run yet).
2. Edit > Variables: add the secrets for your auth mode (section 3). Mark each as **secret**.
3. Project settings > Pipelines > Settings: set **Days to keep runs** to more than the longest gap between two
   scheduled runs (the resume state and the token cache travel as artifacts of the previous run; default 30 days).
4. Run pipeline (manual): choose `authMode`, `environment` (`USGov`), `workspaceNames` (e.g. `Finance*;HR;*Sales*`
   - `*` alone means all workspaces), optionally `stages` = `Inventory` for a quick first test.
5. Open the run, expand **Run ImpactIQ**. With `DeviceCode` the step prints
   `DEVICE CODE SIGN-IN REQUIRED: ... https://microsoft.com/devicelogin ... code XXXXXXXX` - complete it within
   15 minutes (or read it from the Teams/Slack channel if `IMPACTIQ_DEVICECODE_WEBHOOK` is set).
6. When the run finishes, the **Summary** tab shows the ImpactIQ run table (stage, status, items done/failed) and the
   artifacts `impactiq-state`, `impactiq-outputs`, `impactiq-logs` (and `impactiq-backups` when enabled).
7. The schedule (`0 6 * * 1-5` UTC, weekdays, `always: true`) now runs silently: the token cache is restored from the
   previous run's `impactiq-state` artifact and refreshed without a human.

The schedule is UTC and ignores daylight saving. Edit the cron line in `azure-pipelines.yml`; the UI schedule
override in pipeline settings takes precedence if you set one there.

## 3. Secrets and variables

Secret variables are not exposed to scripts automatically; the template maps them with `env:` blocks into the
`IMPACTIQ_*` variables that `ImpactIQ.ps1` reads. Define only the ones your mode needs (missing ones arrive empty and are
ignored):

| Variable (secret) | Mode | Notes |
|---|---|---|
| `IMPACTIQ_TOKEN_CACHE_KEY` | DeviceCode on **hosted** agents | any random string >= 32 chars; encrypts `State\auth\token-cache.json` (AES-256) so it can travel in the `impactiq-state` artifact. Not needed on self-hosted Windows agents (DPAPI) |
| `IMPACTIQ_DEVICECODE_WEBHOOK` | DeviceCode | optional Teams/Slack incoming webhook URL; receives `{ "text": "<device code message>" }` |
| `IMPACTIQ_USERNAME`, `IMPACTIQ_PASSWORD` | Credential | MFA-/CA-exempt account only (see Auth-Options.md 4.2) |
| `IMPACTIQ_PBI_TOKEN`, `IMPACTIQ_FABRIC_TOKEN` | AccessToken | one-off tests only; no refresh |
| `IMPACTIQ_TENANT_ID` | any | tenant GUID; recommended for guest accounts or when `organizations` resolves to the wrong tenant |
| `IMPACTIQ_CLIENT_ID` | DeviceCode / Credential | override the default public client id (rarely needed) |

Put them either directly on the pipeline (Edit > Variables, lock icon) or in a **variable group** (Pipelines >
Library) and pass its name in the `variableGroup` parameter; authorize the pipeline on the group's *Pipeline
permissions* tab. Never write them into the YAML. Masking is best-effort - the tool additionally registers every token
it mints with `##vso[task.setsecret]` and redacts `Password=`/`Bearer` patterns in its own log.

`System.AccessToken` is only mapped in the optional "Commit workbooks" step.

## 4. Hosted vs self-hosted agents

| | Microsoft-hosted `windows-latest` | Self-hosted Windows agent |
|---|---|---|
| Job time limit | 60 min free / **360 min** with a paid parallel job (set `timeoutInMinutes: 360` explicitly - the pipeline does) | none (`timeoutInMinutes: 0`) |
| State between runs | only via artifacts (`restoreState: true`, default) | the agent workspace persists (`$(Build.SourcesDirectory)\State` survives) - set `restoreState: false` so an older artifact never overwrites newer local state |
| Token cache | AES with `IMPACTIQ_TOKEN_CACHE_KEY` (published inside `impactiq-state`) | DPAPI under the agent service account; nothing published |
| Tabular Editor 2 / pbi-tools | committed copies in `Config\`, updated from GitHub each run unless `skipToolUpdate: true`; WER dialogs disabled by the tool | same; downloads can be blocked -> `skipToolUpdate: true` |
| Modules | ImportExcel installed each run (`installModules` parameter, PSGallery must be reachable); Az preinstalled; MicrosoftPowerBIMgmt not needed for DeviceCode/Credential | install once (`Install-Module ImportExcel -Scope CurrentUser` **as the agent account**) and set `installModules: ''` |
| Network | public Azure IPs; Conditional Access "named location" policies and "block public internet access" tenant settings break it | inside your network |
| Compliance | Azure DevOps Services is not FedRAMP-authorized; some agencies forbid hosted agents processing GCC data | your box, your rules |
| Concurrency | fresh VM per job | one agent = runs queue instead of overlapping |

### 4.1 Self-hosted agent quick setup

```powershell
# On the Windows box (Server 2016+ or Windows 10/11), as an administrator:
# Project settings > Agent pools > Add pool "ImpactIQ" (self-hosted) > New agent > download the zip
mkdir C:\azagent; cd C:\azagent; Expand-Archive ~\Downloads\vsts-agent-win-x64-*.zip .
.\config.cmd --unattended --url https://dev.azure.com/<org> --auth pat --token <one-time PAT with Agent Pools (read, manage)> `
  --pool ImpactIQ --agent $env:COMPUTERNAME --runAsService --windowsLogonAccount 'DOMAIN\svc-impactiq' --windowsLogonPassword '<pwd>'
# The service account "svc-impactiq" is the identity that will own the DPAPI token cache and the ImportExcel install.
```

Then in `azure-pipelines.yml` (or the run-pipeline dialog): `pool: { name: ImpactIQ }`, `timeoutInMinutes: 0`,
`restoreState: false`. Bootstrap the sign-in **through a manual pipeline run on that agent** (it runs as
`svc-impactiq`, so the DPAPI cache is created for the right account); complete the device code from the run log.

If the agent runs under `NETWORK SERVICE` (the installer default), the cache lives under
`C:\Windows\ServiceProfiles\NetworkService` - it works, but you cannot bootstrap it from a desktop session; use the
manual-run approach.

## 5. How a run flows (state, artifacts, resume)

```
checkout  ->  restore impactiq-state (previous run)  ->  install modules  ->  Run ImpactIQ  ->  stage + publish artifacts
                 State\runs\<RunId>\...                                      exit 0 / 2 / 1       impactiq-state   (State\runs + State\auth*)
                 State\auth\token-cache.json*                                                      impactiq-outputs (4 workbooks)
                                                                                                    impactiq-logs    (Logs\ImpactIQ_*.log)
                                                                                                    impactiq-backups (optional)
```
`*` only when `IMPACTIQ_TOKEN_CACHE_KEY` is set.

* **Restore** uses `DownloadPipelineArtifact@2` with `buildType: specific`, `definition: $(System.DefinitionId)`,
  `buildVersionToDownload: latestFromBranch`, `allowPartiallySucceededBuilds` + `allowFailedBuilds: true`,
  `continueOnError: true` (the very first run logs `No builds currently exist in the build definition supplied` - harmless).
  Runs that were **cancelled** by the job timeout are not returned by the task; that is why every publish step has
  `condition: always()` and the job has `cancelTimeoutInMinutes: 15` - the state is captured even when the job is
  killed at the 360-minute mark, and the following run resumes it.
* **Resume**: `ImpactIQ.ps1 -Resume Auto` (the template's default) resumes today's run if its manifest is not
  `Completed`, otherwise the newest run of the last 3 days with status `Running`, `Failed` or `CompletedWithErrors`,
  otherwise starts a fresh run for today. Completed stages are skipped (except `Assemble`, always rebuilt), completed
  items are skipped inside stages, failed items are retried. Details in Headless-and-Resume.md section 4.
* **Exit codes** (from the "Run ImpactIQ" step): `0` -> Succeeded; `2` -> `##vso[task.complete
  result=SucceededWithIssues;]` (orange run; check the Failures sheet / `manifest.json`); `1` -> Failed (auth, scope,
  inventory or an unhandled error - the log tells which). A failed run still publishes its state, so the next run
  continues where it stopped.
* **Summary**: the "Stage artifacts" step writes `impactiq-summary.md` from `manifest.json` and uploads it with
  `##vso[task.uploadsummary]` - the run's Summary tab shows stages, counts and the first 50 failures.
* **Backups** (`publishBackups: true`) publishes only backup date-folders written in the last 3 days, so a persistent
  self-hosted workspace does not re-upload months of `.pbix`/`.bim` files every night.
* **Big tenants on hosted agents**: a run that needs more than 6 hours simply spans several scheduled runs (each
  resumes the previous). If you need it in one go, use a self-hosted agent.

## 6. Getting the outputs into Power BI

`Power BI Governance Model.pbit` reads the four workbooks through the parameters `UseWeb` (true/false),
`Base Directory`, `Base Model File`, `Base Report File`, `Base Environment File`, `Base Dataflow File`. With
`UseWeb = false` it opens local files (`C:\Power BI Backups` by default) - fine on a workstation, useless in the Power
BI Service. Two ways to make the Service refresh from what the pipeline produced:

### 6.1 Azure Repos + PAT (no gateway, no service principal) - `commitOutputs: true`

The template's optional step pushes `outputs\*.xlsx` (+ the newest `manifest.json`) to an unprotected branch
(`outputsBranch`, default `data`) with the job's own identity. One-time setup:

1. Project settings > Repos > Repositories > *this repo* > Security > **`<Project> Build Service (<Org>)`** >
   allow **Contribute** and **Create branch**. (If your organization runs with "Limit job authorization scope" off, the
   identity is `Project Collection Build Service (<Org>)` instead.) The identity appears only after the pipeline ran once.
2. Do not put branch policies on the outputs branch. The commit message carries `[skip ci]`.
3. Create an **org-scoped** PAT (User settings > Personal access tokens) with scope **Code (Read)** only. Global
   (all-organization) PATs stop working on 1 Dec 2026 - use org-scoped; note the expiry and rotate it.

Power BI Desktop: change the `Base Directory` query (or add a new parameter set) so the four `Base ...` queries use
`Web.Contents` against the Git *Items* API. Keep the **first argument static** (`https://dev.azure.com/{org}/`) and put
everything else into `RelativePath` / `Query` - that is the documented exception that lets the Service refresh a
dynamic web source without a gateway:

```powerquery-m
// Parameters: AdoOrg = "contoso", AdoProject = "Governance", AdoRepo = "ImpactIQ", AdoBranch = "data"
let
    GetWorkbook = (fileName as text) as binary =>
        Web.Contents(
            "https://dev.azure.com/" & AdoOrg & "/",
            [
                RelativePath = AdoProject & "/_apis/git/repositories/" & AdoRepo & "/items",
                Query = [
                    path = "/outputs/" & fileName,
                    #"versionDescriptor.version" = AdoBranch,
                    #"versionDescriptor.versionType" = "branch",
                    download = "true",
                    #"api-version" = "7.1"
                ]
            ]),
    // one call per workbook; the existing "Base ..." queries then navigate to their sheets
    EnvironmentWorkbook = Excel.Workbook(GetWorkbook("Power BI Environment Detail.xlsx"), null, true),
    ReportWorkbook      = Excel.Workbook(GetWorkbook("Report Detail.xlsx"), null, true),
    ModelWorkbook       = Excel.Workbook(GetWorkbook("Model Detail.xlsx"), null, true),
    DataflowWorkbook    = Excel.Workbook(GetWorkbook("Dataflow Detail.xlsx"), null, true),
    // example: the Workspaces sheet
    Workspaces = EnvironmentWorkbook{[Item = "Workspaces", Kind = "Sheet"]}[Data]
in
    Workspaces
```

Credentials (Desktop: File > Options > Data source settings; Service: dataset Settings > Data source credentials):
data source `https://dev.azure.com/{org}/`, authentication **Basic**, user name blank (or anything), password = the
PAT, privacy level Organizational. If the credential test fails on the base URL, tick **Skip test connection**.
Do **not** use the `Headers = [Authorization = "Basic ..."]` variant with Anonymous credentials - it hard-codes the PAT
inside the model.

Refresh in the Service: the dataset is a cloud source -> no gateway. Schedule refresh after the pipeline's usual end
time. Rotate the PAT before it expires (30-90 days depending on org policy) - the refresh fails with 401 when it lapses.

### 6.2 SharePoint / OneDrive with the template's `UseWeb`

Set `UseWeb = true` and `Base Directory` to the library folder URL, e.g.
`https://agency.sharepoint.us/sites/BI/Shared%20Documents/ImpactIQ/` (GCC SharePoint lives on `sharepoint.us`;
commercial on `sharepoint.com`). The template then builds `Web.Contents(<Base Directory> & <file>)` with
`Dataflow Detail.xlsx` URL-escaped; authenticate the `https://agency.sharepoint.us/` data source with **Organizational
account**. Power BI refreshes this without a gateway. How the files get there is the hard part without a service
principal - see section 7.

### 6.3 Pipeline artifacts only

If neither is possible, download `impactiq-outputs` from the run (or with `az pipelines runs artifact download`) to
the folder that the local `.pbit` reads. No Service refresh in that case.

## 7. SharePoint / OneDrive upload without a service principal

* **OneDrive sync folder** (`sharePointSyncPath` parameter): only on a self-hosted agent configured as an
  **interactive process with auto-logon** (the OneDrive sync client is a per-user UI app; it does not run under a
  service account and there is no client on hosted agents). The step copies the four workbooks into the synced folder;
  the sync client uploads asynchronously.
* **PnP.PowerShell `Connect-PnPOnline -Credentials`** is ROPC (dies with MFA) and since Sept 2024 requires *your own*
  Entra app registration (`Register-PnPEntraIDApp`) and PowerShell 7.4+. Legacy username/password (IDCRL) auth was
  retired on 1 May 2026 and ACS app-only on 2 Apr 2026. If you can register an app with a certificate and
  `Sites.Selected`, do that - it is a service principal, but the least-privileged one possible. Not wired into the
  template on purpose.
* Practical recommendation: use **6.1 (Azure Repos + PAT)** - it needs no upload step at all.

## 8. Pipeline parameters (Run pipeline dialog)

| Parameter | Default | Meaning |
|---|---|---|
| `pool` | `vmImage: windows-latest` | `{ name: <self-hosted pool> }` for self-hosted |
| `authMode` | `DeviceCode` | `DeviceCode` / `Credential` / `AzContext` (`AccessToken` and `Interactive` are not offered - use `extraArgs` for tests) |
| `environment` | `USGov` | `Public`, `USGov`, `USGovHigh`, `USGovMil`, `China`, `Germany` |
| `runMode` | `Workspaces` | `Workspaces` / `Reports` (`extraArgs: -ReportId a,b`) / `Models` (`extraArgs: -DatasetId a,b`) |
| `workspaceNames` | `*` | semicolon-separated, wildcards (`-like`): `Finance*;HR;*Sales*`; `*` = `-AllWorkspaces` |
| `stages` | `` (all) | comma-separated subset of `Inventory,ModelBackup,ReportBackup,ReportDetail,ModelDetail,Dataflows,Extras,Assemble` |
| `extraArgs` | `` | appended verbatim, e.g. `-IncludeMyWorkspace -IncludeUsageMetrics -IncludeAdminApis -ActivityDays 7 -MaxParallelExtracts 3 -ModelDetailMethod Dax -Force` |
| `publishBackups` | `false` | publish `impactiq-backups` (recent `Model/Report/Dataflow Backups` date folders) |
| `commitOutputs` / `outputsBranch` | `false` / `data` | section 6.1 |
| `sharePointSyncPath` | `` | section 7 |
| `variableGroup` | `` | Library variable group holding the `IMPACTIQ_*` secrets |
| `timeoutInMinutes` | `360` | `0` = unlimited (self-hosted) |
| `restoreState` | `true` | `false` on self-hosted agents whose workspace persists |
| `skipToolUpdate` | `false` | do not contact GitHub for Tabular Editor / pbi-tools updates |
| `usePwsh` | `false` | run under PowerShell 7 instead of Windows PowerShell 5.1 |

`ImpactIQ.ps1` itself always receives `-NonInteractive -Resume Auto -BaseFolder $(Build.SourcesDirectory)` from the
template. The exact command line is printed at the top of the "Run ImpactIQ" step (secrets are not part of it).

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Job cancelled at 60:00 | free hosted tier. Buy one parallel job **and** keep `timeoutInMinutes: 360` in the YAML (without it the 60-minute default still applies), or go self-hosted. The next run resumes. |
| `No builds currently exist in the build definition supplied` on the restore step | first run; harmless (`continueOnError`). |
| Every run asks for a device code again | `IMPACTIQ_TOKEN_CACHE_KEY` not set (hosted agents) or changed; artifact retention shorter than the schedule gap; `restoreState: false` on a hosted agent; the refresh token was revoked (password change, CA). Look for `Token cache included` in the "Stage artifacts" log of the previous run. |
| `AADSTS50076` / `53003` in Credential mode | MFA / Conditional Access - switch to DeviceCode (Auth-Options.md). |
| `Install-Module ImportExcel` fails | PSGallery unreachable from the agent - install it once manually as the agent account and set `installModules: ''`; on hosted agents check the NuGet provider bootstrap lines in the log. |
| Run ends with `exit 2` (orange) | item failures - open the Summary tab / `Failures` sheet. Common: XMLA read-only capacity setting (model backup), `executeQueries` tenant setting off (DAX model detail), 403 on a workspace, paginated report export blocked by "Download reports". |
| `429` storms in the log | throttling; the tool honours `Retry-After` (up to 8 retries) - reduce `-MaxParallelExtracts`, run fewer workspaces per schedule, or avoid `-IncludeAdminApis` during business hours. |
| Tabular Editor timeouts (`TimedOut` in the Failures sheet) | raise `-ToolTimeoutMinutes` (default 20) via `extraArgs`, lower `-MaxParallelExtracts`; on hosted agents (2 vCPU) keep parallelism at 2. |
| `git push` rejected in the commit step | grant `<Project> Build Service (<Org>)` Contribute + Create branch on the repo; remove policies from the outputs branch. |
| Power BI Service: "dynamic data sources can't be refreshed" | the first `Web.Contents` argument is not static - keep `https://dev.azure.com/{org}/` literal and move everything else into `RelativePath`/`Query` (section 6.1). |
| Power BI refresh 401 after weeks | the PAT expired or the org disabled global PATs - create an org-scoped PAT and update the credential. |
| Hosted agent cannot sign in although the account works from the office | Conditional Access named-location policy - self-hosted agent, or ask for a CA exclusion for the service account. |
| Fabric sheets empty | expected in GCC (no Fabric); elsewhere the account may lack a Fabric token (DeviceCode redeems the refresh token for `<FabricApiPrefix>/.default`; check the Debug line `Fabric token unavailable`). |
