<p align="center">
<img width="338" height="190" alt="impact iq 16x9" src="https://github.com/user-attachments/assets/436670cb-9aec-4ae9-8a5c-ace341ff9283" />
  <br/>
<em><strong>Be smarter with every change </strong></em>
    <br/>
  <br/>
  <em>One-Click, Designed for Everyone</em>
</p>

# Impact Analysis and Governance for Power BI + Fabric

ImpactIQ backs up every model, report and dataflow you can reach, extracts their metadata down to the visual level,
and loads it all into the **Power BI Governance Model** so you can see what depends on what before you change it.

**v3** keeps the one-click, interactive experience of v2 and adds what an organisation needs to run it unattended:
a headless entry point, checkpoints that let an interrupted run resume where it stopped, sign-in options that work on
a schedule without a service principal, an Azure DevOps pipeline, more data in the same workbooks, and endpoint
tables for the commercial, GCC, GCC High, DoD and China clouds.

### This all-in-one solution is designed to be run by ANYONE.
- Everything within the script is limited to your access within the Power BI and/or Fabric environment.
- All computer requirements are at the user level and do not require admin privileges.
- No app registration, no service principal, no Azure VM. A Pro or PPU licensed user account is all it signs in as.

*Have specific Reports and/or Models downloaded you want to analyze? Don't have direct access to the Workspace but have the PBIX? Check out Impact IQ's local edition here: https://github.com/BeSmarterWithData/ImpactIQ-Local*

**Other Versions:** *[Tenant Admin Edition](https://github.com/BeSmarterWithData/ImpactIQ-TenantAdmin), [Semantic Link Labs Edition](https://github.com/BeSmarterWithData/ImpactIQ-SemanticLinkLabs), [Service Principal Edition](https://github.com/BeSmarterWithData/ImpactIQ-ServicePrincipal)*

## Contents

1. [What it does](#what-it-does)
2. [What changed in v3](#what-changed-in-v3)
3. [Quick start (interactive, 5 minutes)](#-quick-start-interactive-5-minutes)
4. [Getting started: full rollout walkthroughs](#-getting-started-full-rollout-walkthroughs)
   - [Before you start (every option)](#before-you-start-every-option)
   - [Option A: interactive runs on a workstation](#option-a-interactive-runs-on-a-workstation)
   - [Option B: Windows Task Scheduler (no Azure DevOps)](#option-b-windows-task-scheduler-no-azure-devops)
   - [Option C: Azure DevOps with a self-hosted Windows agent (recommended)](#option-c-azure-devops-with-a-self-hosted-windows-agent-recommended)
   - [Option D: Azure DevOps with a Microsoft-hosted agent and device-code sign-in](#option-d-azure-devops-with-a-microsoft-hosted-agent-and-device-code-sign-in)
   - [Option E: Azure DevOps with a Microsoft-hosted agent and an MFA-exempt account](#option-e-azure-devops-with-a-microsoft-hosted-agent-and-an-mfa-exempt-account)
   - [Connecting the Power BI Governance Model](#connecting-the-power-bi-governance-model)
   - [Running it week after week](#running-it-week-after-week)
5. [Where the files go](#-where-the-files-go)
6. [Clouds and endpoints](#-clouds-and-endpoints)
7. [Command-line reference](#-command-line-reference)
8. [Documentation](#-documentation)
9. [Features](#features)

---

## What It Does
This provides a quick and automated way to identify where and how specific fields, measures, and tables are used across Power BI reports in all workspaces down to the visual level. It also backs up and breaks down the details of your models, reports, and dataflows for easy review, giving you an all-in-one **Power BI & Fabric Governance** solution.

### Key Features:
- **Impact Analysis**: Fully understand the downstream impact of data model changes with visual-level lineage, ensuring you don't accidentally break visuals or dashboards, even when multiple reports connect to a model in a different workspace.
- **Used and Unused Objects**: Identify which tables, columns, and measures are actively used and where. Equally as important, see what isn't used and can be safely removed from your model to save space and complexity.
- **Comprehensive Environment Overview**: Gain a clear, detailed view of your entire Power BI environment, including complete breakdowns of your models, reports, and dataflows and their dependencies.
- **Backup Solution**: Automatically backs up every model, report, and dataflow for safekeeping.
- **User-Friendly Output**: the final output is presented in a Power BI Report & Model, making everything easy to explore, analyze, and share with your team.
- **Runs on a schedule**: the same extraction runs headless from Task Scheduler or Azure DevOps, resumes after any interruption and signs in as a user without a service principal.

---

## What changed in v3

Everything v2 produced is still produced, with the same names, so an existing `Power BI Governance Model.pbit`
keeps working. What is new:

| Area | v2 | v3 |
|---|---|---|
| Entry point | one 3,740-line interactive script | `ImpactIQ.ps1` plus fourteen modules under `Config\Modules\`; `Final PS Script.txt` is a thin launcher for the interactive experience |
| Where it runs from | hard-coded `C:\Power BI Backups` | the folder you downloaded it to; backups and workbooks can be moved with `-BackupFolder` / `-OutputFolder` |
| Headless | not possible (dialogs, browser sign-in, `Read-Host`) | `-NonInteractive` plus a parameter or `IMPACTIQ_*` variable for every former prompt; no scope given means the run stops instead of scanning the tenant |
| Interruptions | start over | every workspace, model, report, model-detail CSV and dataflow is checkpointed the moment it finishes; re-running skips what succeeded and retries what failed; `-TimeBudgetMinutes` pauses cleanly before a job cap and resumes on the next run |
| Sign-in | browser every 55 minutes | silent token refresh; `DeviceCode` with an encrypted refresh-token cache, `Credential`, `AzContext`, `AccessToken`, and the original `Interactive` |
| Robustness | silent failures, no retries | one HTTP wrapper with retry and back-off (429 `Retry-After`, 5xx, network), timeouts and captured output for Tabular Editor and pbi-tools, per-item failure records, a `Failures` sheet, exit codes `0 / 2 / 3 / 1` |
| Clouds | endpoint table with two wrong hosts | verified tables for Public, GCC, GCC High, DoD, China (and the retired Germany cloud); aliases `GCC`, `GCCHigh`, `DoD`; every REST, OAuth, XMLA, Fabric and portal URL follows `-Environment` |
| Fabric absent (GCC) | retries and errors per call | detected once per run (token refused or host unreachable), then skipped without a request; the Fabric-only sheets stay empty and everything else is unaffected |
| Model detail | Tabular Editor 2 over XMLA only | Tabular Editor, or the built-in `.bim` parser, or DAX `INFO.VIEW.*` over `executeQueries` (`-ModelDetailMethod Auto\|TabularEditor\|Bim\|Dax\|Both`) |
| Data | 17 sheets | plus Dashboards, DashboardTiles, Capacities, WorkspaceUsers, DatasetUsers, DatasetParameters, DatasetDQRefreshSchedule, RunSummary, Failures, InventoryErrors, ReportExports; optional `-IncludeUsageMetrics` and `-IncludeAdminApis` |
| Scheduling | none | `pipelines/azure-pipelines.yml` (state restored from the previous run, artifacts, optional commit of the workbooks for a gateway-free refresh) |
| Quality | none | parser, PSScriptAnalyzer (5.1 and 7 compatibility rules) and a 575-test Pester suite; `docs/Validation-Report.md` lists the v2 defects that were fixed |

---

## 🚀 Quick Start (interactive, 5 minutes)

The v2 experience: dialogs, browser sign-in, pickers, and the workbooks next to the script when it finishes.

1. **Download this repository** (Code > Download ZIP, or `git clone`) and extract it into any folder, for example
   `C:\ImpactIQ`. Keep the layout: `ImpactIQ.ps1`, `Final PS Script.txt`, `Config\`, `Power BI Governance Model.pbit`.
   > The one-click batch file published with v2 downloads the repository it was built for into `C:\Power BI Backups`.
   > It still works with v3 (the scripts run from wherever they are), but to get the v3 code, download this
   > repository.
2. **Run the launcher.** Either rename `Final PS Script.txt` to `Final PS Script.ps1` and run it, or open PowerShell in
   the folder and paste the file's contents. PowerShell may offer to install the `ImportExcel` and Power BI modules
   for your user; no admin rights are needed.
3. **Answer the prompts**: environment (`Public` after 60 seconds; `USGov` for GCC), sign in, choose whether to run
   against workspaces, reports or models, then pick them.
4. **Wait.** The console shows each stage and ends with a per-stage summary, the four workbook paths and the log file.
   If the run is interrupted, run the launcher again: finished items are skipped.
5. **Open `Power BI Governance Model.pbit`**, set `Base Directory` to the folder, let it refresh, save as `.pbix`.

📂 **All backups and the four workbooks land in the folder you ran it from.** To put them elsewhere, set
`IMPACTIQ_BACKUP_FOLDER` and/or `IMPACTIQ_OUTPUT_FOLDER` before running the launcher (or pass `-BackupFolder` /
`-OutputFolder` to `ImpactIQ.ps1`).

---

## 🧭 Getting started: full rollout walkthroughs

Pick the option that matches where the run can live and how the identity can be kept alive between runs. The
decision in one line: **any Windows machine that stays on → Option C** (or B without Azure DevOps); **no machine at
all → Option D**; **an MFA-exempt service account that security accepts → Option E**.

| | Runs where | Signs in how | Time cap | Needs | Best for |
|---|---|---|---|---|---|
| **A** | your workstation, by hand | browser | none | nothing extra | a single owner, on demand |
| **B** | a Windows box, Task Scheduler | device code once, DPAPI cache | none | a box that stays on | one owner, nightly, no Azure DevOps |
| **C** | Azure DevOps, self-hosted agent on a Windows box | device code once, DPAPI cache (or AzContext) | none | the free agent installed as a service | **teams; the recommended default** |
| **D** | Azure DevOps, Microsoft-hosted `windows-latest` | device code once, AES-encrypted cache carried between runs | 60 min free / 360 min with one paid parallel job | one secret variable; a paid parallel job for big tenants | no Windows machine available |
| **E** | Azure DevOps, Microsoft-hosted | username and password (ROPC) | same as D | an MFA- and Conditional-Access-exempt, cloud-only account | only with security sign-off |

### Before you start (every option)

1. **An account to run as.** Pro or PPU license. Workspace **Viewer** gives the inventory; **Contributor** (or
   Member/Admin) is needed for model backups over XMLA, report downloads, DAX expressions and dataflow contents.
   Give it access to every workspace you want documented. Keep it out of the Fabric administrator role unless you
   want `-IncludeAdminApis`.
2. **Tenant and capacity settings** (Power BI admin portal): *Download reports* on (report backups), *Allow XMLA
   endpoints* on and the capacity's *XMLA Endpoint* set to Read or Read Write (model backups on dedicated capacity),
   *Semantic Model Execute Queries REST API* on (the DAX fallback and usage metrics). Missing settings do not stop
   the run; the affected items are recorded in the `Failures` sheet.
3. **The machine** (Options A, B, C): Windows 10/11 or Server 2016+, Windows PowerShell 5.1 (built in) or
   PowerShell 7. Tabular Editor 2 and pbi-tools are included under `Config\` and need no installation; Power BI
   Desktop on the machine lets pbi-tools read models embedded in Pro-workspace PBIX files. Outbound HTTPS to the
   Power BI endpoints of your cloud and, for tool updates, to `github.com` (or use `-SkipToolUpdate`).
4. **The cloud.** Decide the `-Environment` value: `Public` (commercial), `USGov` (GCC), `USGovHigh` (GCC High),
   `USGovMil` (DoD), `China`. Aliases `GCC`, `GCCHigh`, `DoD` work too. Every URL the tool uses follows this value
   and the resolved hosts are printed at the top of each run (see [Clouds and endpoints](#-clouds-and-endpoints)).
5. **Get the files.** Download this repository into one folder on the machine that will run it (or into the Azure
   Repos project for Options C, D, E). That folder is the *base folder*: it holds `Config\`, `State\`, `Logs\` and,
   by default, the backups and the workbooks.
6. **Do one interactive run first** (Option A) against a small workspace. It confirms the account, the settings and
   the tools before anything is scheduled, and it shows you what the outputs look like.

### Option A: interactive runs on a workstation

1. Download the repository into a folder, for example `C:\ImpactIQ`.
2. Run `Final PS Script.txt` (rename to `.ps1`, or paste it into a PowerShell window opened in that folder).
   The launcher finds `ImpactIQ.ps1` next to itself.
3. Choose the environment, sign in in the browser, choose the run mode and the workspaces, reports or models.
4. When it finishes, open `Power BI Governance Model.pbit`, point `Base Directory` at `C:\ImpactIQ`, refresh, save as
   `.pbix`.
5. Re-run whenever you want fresh data. A run interrupted the same day resumes; `-Force` on `ImpactIQ.ps1` starts
   the day over.

Want the same without the launcher? `ImpactIQ.ps1 -AuthMode Interactive -Environment USGov` opens the same dialogs;
add `-WorkspaceName "Finance*"` to skip the picker.

### Option B: Windows Task Scheduler (no Azure DevOps)

Runs nightly on a workstation or a server as a user who signed in once. Nothing leaves the machine.

1. Download the repository into a folder the scheduled user can write to, for example `C:\ImpactIQ`.
2. Decide the scheduled identity: a service account with the licenses and roles from *Before you start*, or your
   own account. **The bootstrap in step 3 must run as that identity**: the token cache is protected with Windows
   DPAPI for that user on that machine.
3. **Bootstrap the sign-in once**, in a PowerShell window running as the scheduled user (use `runas` for a service
   account):
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "C:\ImpactIQ\ImpactIQ.ps1" -NonInteractive -Environment USGov -AuthMode DeviceCode -Stages Inventory -WorkspaceName "Finance"
   ```
   The console prints `DEVICE CODE SIGN-IN REQUIRED ... https://microsoft.com/devicelogin ... code XXXXXXXXX`.
   Complete it within 15 minutes. The refresh token is now cached in `State\auth\token-cache.json`.
4. Run the same command again: it must **not** ask for a code. If it does, the bootstrap ran as a different user.
5. **Create the task** (Task Scheduler > Create Task): run as the same user, *Run whether user is logged on or not*,
   *Do not store password* unchecked, trigger daily at a quiet hour. Action:
   ```text
   Program:   powershell.exe
   Arguments: -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\ImpactIQ\ImpactIQ.ps1" -NonInteractive -Environment USGov -AuthMode DeviceCode -AllWorkspaces -IncludeMyWorkspace
   Start in:  C:\ImpactIQ
   ```
   Add `-BackupFolder "D:\PBI Backups"` or `-OutputFolder "\\server\share\ImpactIQ"` to move the data; add
   `-ModelDetailMethod Dax` if the machine cannot run Tabular Editor.
6. **Check the result**: the task's *Last Run Result* is the exit code (`0` ok, `2` finished with item failures,
   `3` paused by a time budget, `1` fatal). Details are in `Logs\ImpactIQ_<timestamp>.log` and the `Failures` sheet.
7. **Connect the template**: open the `.pbit` on the machine (or on any machine that can read the output folder),
   set `Base Directory`, refresh, publish. For a scheduled refresh in the Service, see
   [Connecting the Power BI Governance Model](#connecting-the-power-bi-governance-model).
8. **Every ~90 days** (or after a password change or a Conditional Access change) the log says `invalid_grant`:
   repeat step 3.

### Option C: Azure DevOps with a self-hosted Windows agent (recommended)

Any domain-joined Windows machine that stays on can host the free self-hosted agent. You get run history,
artifacts, a Summary tab with the stage table, no job time cap, native Tabular Editor and pbi-tools, and no secret
stored in Azure DevOps.

1. **Put the repository in Azure Repos** (import this repository or push a clone). The repository root is the base
   folder; the pipeline passes it as `-BaseFolder $(Build.SourcesDirectory)`.
2. **Create the agent pool and install the agent** on the Windows box, as an administrator:
   ```powershell
   # Project settings > Agent pools > Add pool "ImpactIQ" (self-hosted) > New agent > download the zip
   mkdir C:\azagent; cd C:\azagent; Expand-Archive ~\Downloads\vsts-agent-win-x64-*.zip .
   .\config.cmd --unattended --url https://dev.azure.com/<org> --auth pat --token <one-time PAT with Agent Pools (read, manage)> `
     --pool ImpactIQ --agent $env:COMPUTERNAME --runAsService --windowsLogonAccount 'DOMAIN\svc-impactiq' --windowsLogonPassword '<pwd>'
   ```
   `svc-impactiq` is the account from *Before you start*; it owns the DPAPI token cache and the module install.
3. **Install ImportExcel once as that account** (`Install-Module ImportExcel -Scope CurrentUser`) if the box cannot
   reach the PowerShell Gallery during runs; otherwise the pipeline installs it.
4. **Create the pipeline**: Pipelines > New pipeline > Azure Repos Git > this repo > *Existing Azure Pipelines YAML
   file* > `/pipelines/azure-pipelines.yml` > Save. In the YAML (or the run dialog) set
   `pool: { name: ImpactIQ }`, `timeoutInMinutes: 0`, `restoreState: false` (the agent workspace persists, so an
   older artifact must never overwrite it).
5. **Bootstrap the sign-in through a manual run**: Run pipeline > `authMode` `DeviceCode`, `environment` `USGov`,
   `workspaceNames` a small pattern such as `Finance*`, `stages` `Inventory,Assemble`. Open the run, expand
   **Run ImpactIQ**, complete the device code from the log within 15 minutes. Because the job runs as
   `svc-impactiq`, the DPAPI cache is created for the right account. Optional: store a Teams/Slack incoming webhook
   as the secret variable `IMPACTIQ_DEVICECODE_WEBHOOK` and the code is posted there too.
6. **Run it once more by hand** with the full scope (`workspaceNames` `*`, `stages` empty, `extraArgs`
   `-IncludeMyWorkspace`). It must not ask for a code. Check the Summary tab and the `impactiq-outputs` artifact.
7. **Turn on the schedule.** The YAML schedules weekdays at 06:00 UTC with `always: true`; edit the cron line to
   taste. Keep the interval longer than a run, or keep one agent in the pool so a second run queues.
8. **Get the workbooks into Power BI**: set `commitOutputs: true` so each run pushes the four workbooks to the
   `data` branch, then follow [Connecting the Power BI Governance Model](#connecting-the-power-bi-governance-model).
   On a self-hosted agent you can instead copy them to a folder the template reads (`sharePointSyncPath`, or
   `-OutputFolder` through `extraArgs`).
9. **Re-bootstrap** (step 5) when the log says `invalid_grant`: roughly every 90 idle days, after a password change
   or a Conditional Access change.

Prefer an Az PowerShell login on the box? Run `Connect-AzAccount -UseDeviceAuthentication` and
`Enable-AzContextAutosave -Scope CurrentUser` once as the agent account and use `authMode` `AzContext`
(GCC uses the default `AzureCloud` environment; GCC High and DoD need `-Environment AzureUSGovernment`).

### Option D: Azure DevOps with a Microsoft-hosted agent and device-code sign-in

Zero infrastructure. A human completes the device code once; the refresh token is encrypted with a key you keep in
Azure DevOps and travels between runs inside the `impactiq-state` artifact.

1. **Put the repository in Azure Repos** (as in Option C step 1).
2. **Buy one Microsoft-hosted parallel job** (Organization settings > Parallel jobs) unless the tenant is small.
   The free tier caps a job at 60 minutes; a paid job allows 360. ImpactIQ chains runs beyond that, but every run
   must end on its own, so `timeBudgetMinutes` is required: `55` on the free tier, `350` with a paid job.
3. **Create the cache key** and store it as a **secret** pipeline variable `IMPACTIQ_TOKEN_CACHE_KEY`
   (Edit > Variables, lock icon), or in a variable group named in the `variableGroup` parameter:
   ```powershell
   [Convert]::ToBase64String((1..48 | ForEach-Object { Get-Random -Maximum 256 }) -as [byte[]])
   ```
   Optional secret `IMPACTIQ_DEVICECODE_WEBHOOK` (Teams/Slack incoming webhook) and `IMPACTIQ_TENANT_ID` for guest
   accounts.
4. **Set run retention** (Project settings > Pipelines > Settings > *Days to keep runs*) to more than the longest gap
   between two scheduled runs: the state and the token cache are restored from the previous run's artifact.
5. **Create the pipeline** from `/pipelines/azure-pipelines.yml` (Option C step 4) and keep the default
   `pool: vmImage: windows-latest`, `timeoutInMinutes: 360`, `restoreState: true`.
6. **Bootstrap the sign-in through a manual run**: `authMode` `DeviceCode`, `environment` `USGov`,
   `timeBudgetMinutes` `55` or `350`, `workspaceNames` a small pattern, `stages` `Inventory,Assemble`. Complete the
   device code from the **Run ImpactIQ** step log (or the webhook message) within 15 minutes. The run publishes the
   encrypted cache inside `impactiq-state`; the *Stage artifacts* log says `Token cache included`.
7. **Run once more by hand** with the full scope; it must not ask for a code. A large tenant ends with exit `3`
   (*Paused*, orange run) and partial workbooks; the next run resumes it.
8. **Turn on the schedule** (Option C step 7). Each scheduled run restores the previous state, refreshes the token
   silently, continues or starts a fresh day, and publishes `impactiq-state`, `impactiq-outputs`, `impactiq-logs`
   (and `impactiq-backups` when `publishBackups: true`).
9. **Get the workbooks into Power BI** with `commitOutputs: true` and the Azure Repos pattern below. Hosted agents
   have no OneDrive client, so SharePoint sync is not an option here.
10. **Watch for** `invalid_grant` / `new device-code flow` in the log: repeat step 6. Watch for runs *Canceled* at
    the cap: that means `timeBudgetMinutes` is missing or too high, and the next run cannot restore the cancelled
    run's state.

Things that break this option: a Conditional Access named-location policy (hosted agents use public Azure IP
ranges), a tenant setting that blocks public internet access, and agency rules that forbid hosted agents processing
GCC data. In those cases use Option C.

### Option E: Azure DevOps with a Microsoft-hosted agent and an MFA-exempt account

Fully unattended with nothing to re-bootstrap, at the price of an account that is exempt from MFA and Conditional
Access. Microsoft is retiring the password grant; expect this to stop working one day.

1. Steps 1, 2, 4 and 5 of Option D.
2. Create a **cloud-only** service account (password-hash sync or pass-through auth; federated accounts cannot use
   the password grant), license it, give it the workspace roles, and have security exclude it from MFA and
   Conditional Access.
3. Store `IMPACTIQ_USERNAME` and `IMPACTIQ_PASSWORD` as secret pipeline variables.
4. Run the pipeline by hand with `authMode` `Credential`. `AADSTS50076`, `50079`, `53003` or `65001` in the log
   mean the account is still MFA- or CA-bound; there is no workaround inside the tool, switch to Option C or D.
5. Schedule it and connect the template as in Option D.

### Connecting the Power BI Governance Model

`Power BI Governance Model.pbit` reads the four workbooks through the parameters `UseWeb`, `Base Directory`,
`Base Model File`, `Base Report File`, `Base Environment File` and `Base Dataflow File`.

- **Local files (Options A, B)**: `UseWeb = false`, `Base Directory` = the output folder. Refresh in Desktop. The
  Service can refresh a local folder only through an on-premises data gateway installed on that machine.
- **Azure Repos with a read-only PAT (Options C, D, E; no gateway)**: set `commitOutputs: true` so the pipeline
  pushes `outputs\*.xlsx` to the `data` branch. One-time: give `<Project> Build Service (<Org>)` *Contribute* and
  *Create branch* on the repo, create an org-scoped PAT with *Code (Read)* only, and change the template's queries
  to `Web.Contents("https://dev.azure.com/<org>/", [RelativePath = ..., Query = ...])` as shown in
  [docs/Azure-DevOps.md](docs/Azure-DevOps.md) section 6.1. The first argument must stay static, that is what lets
  the Service refresh it as a cloud source. Credentials: Basic, password = the PAT. Schedule the dataset refresh
  after the pipeline's usual end time and rotate the PAT before it expires.
- **SharePoint / OneDrive**: `UseWeb = true`, `Base Directory` = the library folder URL (`sharepoint.us` in GCC),
  Organizational account. Getting the files there without a service principal needs a self-hosted agent running
  interactively with the OneDrive client (`sharePointSyncPath`). See [docs/Azure-DevOps.md](docs/Azure-DevOps.md)
  sections 6.2 and 7.

### Running it week after week

- **Exit codes**: `0` clean; `2` finished with item failures (open the `Failures` sheet or the pipeline Summary tab;
  the next run retries them); `3` paused by the time budget (the next run continues); `1` fatal (sign-in, no scope,
  inventory failure; the log says which).
- **Resume rules**: the same day resumes automatically; an unfinished run from the last three days is picked up by
  `-Resume Auto`; `-Force` wipes today's state and backup folders; `-Stages Assemble -Resume Always` rebuilds the
  workbooks from the last run without any API call.
- **Refresh token**: 90 days of inactivity, a password change, an admin *revoke sessions* or a new Conditional
  Access policy end it. The log says `invalid_grant`; redo the bootstrap step of your option.
- **Throttling**: `429` storms mean fewer workspaces per run, `-MaxParallelExtracts 1`, or keeping
  `-IncludeAdminApis` out of business hours; the tool already honours `Retry-After`.
- **Tool updates**: Tabular Editor 2 and pbi-tools are refreshed from GitHub at the start of each run unless
  `-SkipToolUpdate` / `IMPACTIQ_OFFLINE=1`.
- **Two runs at once** against the same base folder are not supported. Keep the schedule interval longer than a run.
- **Windows validation**: the automated tests run on Linux with mocked APIs; `docs/Validation-Report.md` section 4
  lists what to confirm on a Windows box the first time (5.1 runtime, the C# extractor scripts, pbi-tools, DPAPI).

---

## 📂 Where the files go

```
<base folder>  (the folder ImpactIQ.ps1 runs from; -BaseFolder / IMPACTIQ_BASE_FOLDER)
  ImpactIQ.ps1, Final PS Script.txt, Power BI Governance Model.pbit
  Config\                      csx scripts, Blank Model.bim, TabularEditor\, PBI Tools\, Modules\, SheetContract.json
  State\runs\<yyyy-MM-dd>\     manifest.json, per-item checkpoints, inventory JSON, DAX and dataflow extracts, tool logs
  State\auth\token-cache.json  encrypted refresh token (DeviceCode)
  Logs\ImpactIQ_<timestamp>.log
  Model Backups\<yyyy-MM-dd>\    <Workspace> ~ <Model>.bim, .csv          }  -BackupFolder / IMPACTIQ_BACKUP_FOLDER
  Report Backups\<yyyy-MM-dd>\   <Workspace> ~ <Report>.pbix|.rdl, *.txt  }  moves these three
  Dataflow Backups\<yyyy-MM-dd>\ <Workspace> ~ <Dataflow>.txt|.pq         }
  Power BI Environment Detail.xlsx, Report Detail.xlsx,                   }  -OutputFolder / IMPACTIQ_OUTPUT_FOLDER
  Model Detail.xlsx, Dataflow Detail.xlsx                                 }  moves these four
```

`Config\`, `State\` and `Logs\` always stay in the base folder. A relative `-BackupFolder Data` is created under the
base folder; an absolute path or a UNC share works too. Keep the same folders between the runs of one day: the
checkpoints record where each file was written.

---

## 🌐 Clouds and endpoints

`-Environment` (or `IMPACTIQ_ENVIRONMENT`) selects one row; every REST, sign-in, token, XMLA, Fabric and portal URL
comes from it, and the run prints the resolved hosts in its first lines.

| `-Environment` | Portal | REST API | Sign-in authority | XMLA | Fabric API |
|---|---|---|---|---|---|
| `Public` (aliases `Commercial`, `Global`) | app.powerbi.com | api.powerbi.com | login.microsoftonline.com | powerbi://api.powerbi.com | api.fabric.microsoft.com |
| `USGov` (alias `GCC`) | app.powerbigov.us | api.powerbigov.us | login.microsoftonline.com | powerbi://api.powerbigov.us | not offered; detected and skipped |
| `USGovHigh` (alias `GCCHigh`) | app.high.powerbigov.us | api.high.powerbigov.us | login.microsoftonline.us | powerbi://api.high.powerbigov.us | unverified; `-FabricApiPrefixOverride` |
| `USGovMil` (alias `DoD`) | app.mil.powerbigov.us | api.mil.powerbigov.us | login.microsoftonline.us | powerbi://api.mil.powerbigov.us | unverified; `-FabricApiPrefixOverride` |
| `China` | app.powerbi.cn | api.powerbi.cn | login.chinacloudapi.cn | powerbi://api.powerbi.cn | api.fabric.microsoft.cn |
| `Germany` (retired cloud, kept for compatibility) | app.powerbi.de | api.powerbi.de | login.microsoftonline.de | powerbi://api.powerbi.de | api.fabric.microsoft.de |

GCC tenants sign in through commercial Entra ID (`login.microsoftonline.com`); only the Power BI hosts change. GCC High
and DoD use `login.microsoftonline.us` and, for `AzContext`, the `AzureUSGovernment` environment. Where Fabric is not
offered, the first refused token or unreachable host marks Fabric unavailable for the run and the Fabric-only sheets
stay empty; nothing else is affected. Token resources and the full table are in
[docs/Auth-Options.md](docs/Auth-Options.md) section 3.

---

## ⌨️ Command-line reference

The most used parameters of `ImpactIQ.ps1`; every one has an environment-variable twin for schedulers that cannot
pass arguments (full list in [docs/Headless-and-Resume.md](docs/Headless-and-Resume.md)).

| Parameter | Variable | Meaning |
|---|---|---|
| `-BaseFolder` | `IMPACTIQ_BASE_FOLDER` | deployment folder; default: where the script runs from |
| `-BackupFolder`, `-OutputFolder` | `IMPACTIQ_BACKUP_FOLDER`, `IMPACTIQ_OUTPUT_FOLDER` | move the backups / the workbooks; default: the base folder |
| `-Environment` | `IMPACTIQ_ENVIRONMENT` | `Public`, `USGov`/`GCC`, `USGovHigh`/`GCCHigh`, `USGovMil`/`DoD`, `China` |
| `-AuthMode` | (see variables below) | `Auto`, `Interactive`, `DeviceCode`, `Credential`, `AzContext`, `AccessToken` |
| `-TokenCacheKey` | `IMPACTIQ_TOKEN_CACHE_KEY` | AES key for the refresh-token cache on hosted agents (DPAPI on Windows without it) |
| `-DeviceCodeWebhookUrl` | `IMPACTIQ_DEVICECODE_WEBHOOK` | Teams/Slack webhook that receives the device code |
| `-Credential` | `IMPACTIQ_USERNAME`, `IMPACTIQ_PASSWORD` | Credential mode (MFA-exempt accounts only) |
| `-NonInteractive` | automatic under `TF_BUILD` / `CI` | no dialogs, no browser; no scope = stop |
| `-AllWorkspaces`, `-WorkspaceName`, `-WorkspaceId`, `-IncludeMyWorkspace` | | scope in `Workspaces` mode (`-WorkspaceName` takes `-like` wildcards) |
| `-RunMode Reports -ReportId ...`, `-RunMode Models -DatasetId ...` | | scope by report or model; connected objects are added automatically |
| `-Stages`, `-SkipStages` | | `Inventory, ModelBackup, ReportBackup, ReportDetail, ModelDetail, Dataflows, Extras, Assemble` |
| `-Resume Auto\|Always\|Never`, `-Force`, `-RefreshInventory` | | resume rules |
| `-TimeBudgetMinutes` | | stop cleanly N minutes after start, exit `3`, resume next run |
| `-ModelDetailMethod` | | `Auto`, `TabularEditor`, `Bim`, `Dax`, `Both` |
| `-IncludeUsageMetrics`, `-IncludeAdminApis`, `-ActivityDays` | | optional extra sheets |
| `-SkipToolUpdate` | `IMPACTIQ_OFFLINE=1` | no Tabular Editor / pbi-tools downloads |
| `-Verbose` | `IMPACTIQ_DEBUG=1` | echo Debug lines (always in the log file) |

Examples:

```powershell
# everything the account can see, GCC, unattended, resumes an unfinished run automatically
.\ImpactIQ.ps1 -NonInteractive -Environment USGov -AuthMode DeviceCode -AllWorkspaces -IncludeMyWorkspace

# two workspace families, backups on a share, workbooks where the template reads them
.\ImpactIQ.ps1 -NonInteractive -Environment GCC -WorkspaceName 'Finance*','HR' -BackupFolder '\\files\PBI Backups' -OutputFolder 'D:\Governance'

# specific reports (their models and model workspaces are added automatically)
.\ImpactIQ.ps1 -NonInteractive -Environment Public -RunMode Reports -ReportId <guid>,<guid>

# rebuild the four workbooks from the latest run without any API call
.\ImpactIQ.ps1 -NonInteractive -Resume Always -Stages Assemble

# run the quality gate (parser, PSScriptAnalyzer, Pester; no tenant needed)
pwsh -NoProfile -File .\tests\Invoke-Tests.ps1
```

---

## 📚 Documentation

| Page | Read it when |
|---|---|
| [docs/Automation.md](docs/Automation.md) | you want the decision matrix and the five-minute setups behind the walkthroughs above |
| [docs/Auth-Options.md](docs/Auth-Options.md) | you need to decide how a scheduled run signs in, or you are on a GCC / sovereign tenant |
| [docs/Azure-DevOps.md](docs/Azure-DevOps.md) | pipeline parameters, secrets, hosted vs self-hosted, artifacts and resume, `Web.Contents` for the template, troubleshooting |
| [docs/Headless-and-Resume.md](docs/Headless-and-Resume.md) | every parameter and variable, the state layout, resume rules, exit codes, the time budget |
| [docs/Data-Coverage.md](docs/Data-Coverage.md) | which sheet comes from which API, what permission it needs, what the optional flags add, what is empty in GCC |
| [docs/Validation-Report.md](docs/Validation-Report.md) | the v2 audit findings, what v3 fixed, and what still needs a Windows run to confirm |
| [tests/README.md](tests/README.md) | running the Pester suite on Windows PowerShell 5.1 or PowerShell 7 |

---

### ℹ️ Additional Notes

> ⚙️ *PowerShell may prompt to install required modules.*
> No admin access is needed; they install at the user level (`ImportExcel`; `MicrosoftPowerBIMgmt` only for the
> interactive browser sign-in).

> 🧰 *This setup uses the portable version of Tabular Editor 2 (v2.27.2).*
> You don't need it preinstalled. It runs locally from the folder with no differences.
> https://github.com/TabularEditor/TabularEditor _(MIT License)_

> 🧠 *Model backups use XMLA (for PPU, Premium, Fabric).*
> For Pro workspaces, `pbi-tools` extracts the BIM from the PBIX.
> Includes `pbi-tools v1.2`: https://github.com/pbi-tools/pbi-tools _(AGPL 3.0 License)_
> Without either, `-ModelDetailMethod Dax` documents the model over the REST API.

> 🚨 *Using Tabular Editor 3?*
> Tabular Editor 2 is still included and required for this because TE3 doesn't support command line execution.
>
> 🧩 *Model refresh error in Power BI Desktop?*
> If you see:
> _**"Query XXXXXX references other queries or steps..."**_
>
> Update your Power BI Desktop privacy settings:
> **File → Options and settings → Options → Privacy**
> Then select either:
> - "Combine data according to each file's Privacy Level settings"
>   **or**
> - "Always ignore Privacy Level settings"

---
## Features

---

### 1. Workspace and Power BI Environment Metadata Extraction
- Leverages Power BI REST API to gather information about Power BI workspaces, datasets, reports, report pages, apps, dashboards, tiles, users, parameters, refresh schedules and capacities.
- Exports the extracted metadata into a structured Excel workbook with separate worksheets for each entity.
- You must have at least read access within workspaces. 'My Workspace' also included.
- <img width="1255" alt="image" src="https://github.com/user-attachments/assets/515ce3e5-ec56-467a-a421-9da05889eaa5">


### 2. Model Backup and Metadata Extract
- Saves exported models in a structured folder hierarchy based on workspace and dataset names.
- Leverages Tabular Editor 2 and C# to extract the metadata and output within an Excel File; the built-in `.bim` parser or DAX `INFO.VIEW.*` queries take over where Tabular Editor cannot run.
- All backups are saved with the following format: Workspace Name ~ Model Name.
- You must have edit rights on the related model. Works with all Pro, Premium-Per-User, Premium, and Fabric Capacity workspaces. 'My Workspace' also included. Both XMLA and non-XMLA models.
<img width="695" alt="image" src="https://github.com/user-attachments/assets/c3e021b8-6dfe-40c9-bfa5-b9d4471a8fa3">


### 3. Report Backup and Metadata Extract
- Backs up Power BI and Paginated Reports from Power BI workspaces, cleaning report names and determining file types (`.pbix` or `.rdl`) for export.
- Leverages Tabular Editor 2 and C# to extract the Visual Object Layer metadata and output within an Excel File (credit to @m-kovalsky for initial work on this)
- Paginated Reports are only backed up (no metadata extraction).
- All backups are saved with the following format: Workspace Name ~ Report Name.
- You must have edit rights on the related report. Works with all Pro, Premium-Per-User, Premium, and Fabric Capacity workspaces. 'My Workspace' also included.
- <img width="554" alt="image" src="https://github.com/user-attachments/assets/cf88aac7-6f32-445a-96c7-6bc36fcab9aa">


### 4. Dataflow Backup and Metadata Extract
- Extracts dataflows from Power BI workspaces, formatting and organizing their contents, including query details.
- Leverages PowerShell to parse and extract the metadata and output within an Excel File.
- All backups are saved with the following format: Workspace Name ~ Dataflow Name.
- Must have edit rights on the related dataflow. 'Ownership' of the Dataflow is not required. Works with all Pro, Premium Capacity, Fabric Capacity workspaces. 'My Workspace' also included.
- <img width="542" alt="image" src="https://github.com/user-attachments/assets/67e83016-4bc7-4cf5-8d94-1a9779aad6d8">

### 5. Model Connection Details Metadata Extract
- Leverages Power BI REST API to gather all model connection details.
- Exports the extracted metadata into the same structured excel workbook as the Power BI Environment Information Extract
- You must have read permissions on the related model.

### 6. Model Refresh History Metadata Extract
- Leverages Power BI REST API to gather all model refresh history (limited to the same history shown in the Service).
- Exports the extracted metadata into the same structured excel workbook as the Power BI Environment Detail Extract
- You must have read permissions on the related model.

### 7. Model Refresh Schedule Metadata Extract
- Leverages Power BI REST API to gather all model refresh schedule settings including enabled status, time zone, schedule days, and times (import and DirectQuery schedules).
- Exports the extracted metadata into the same structured excel workbook as the Power BI Environment Detail Extract
- You must have read permissions on the related model.

### 8. Dataflow Connection Details Metadata Extract
- Leverages Power BI REST API to gather all Dataflow connection details.
- Exports the extracted metadata into the same structured excel workbook as the Power BI Environment Detail Extract
- You must have read permissions on the related Dataflow.

### 9. Dataflow Refresh History Metadata Extract
- Leverages Power BI REST API to gather all Dataflow refresh history (limited to the same history shown in the Service).
- Exports the extracted metadata into the same structured excel workbook as the Power BI Environment Detail Extract
- You must have read permissions on the related Dataflow.

### 10. Power BI Governance Model
- Combines extracts into a Semantic Model to allow easy exploring, impact analysis, and governance of all Power BI Reports, Models, and Dataflows across all Workspaces
- Works for anyone who runs the script and has at least 1 model and report. Dataflow not required.
- Public example (limited due to no filter pane): https://app.powerbi.com/view?r=eyJrIjoiNmMxYWQ2ZTItZDM4ZS00MGM1LTlhMDQtN2I1OTMwMzI0OTg2IiwidCI6ImUyY2Y4N2QyLTYxMjktNGExYS1iZTczLTEzOGQyY2Y5OGJlMiJ9

## Special Notes
- Tokens are refreshed silently before they expire; the v2 rule of signing in again every 55 minutes is gone.
- The run defaults to what you select. Headless, `-AllWorkspaces` scans everything; `-WorkspaceName`, `-WorkspaceId`, `-ReportId` or `-DatasetId` narrow it; with nothing selected a headless run stops rather than scanning the tenant by accident.
- For the best user experience, the final Power BI Governance Model output is **from the perspective of the Report**. This means that when looking at a Workspace where Reports have the Model sitting in a different Workspace (i.e. multiple reports connected to a model in a different workspace), the Model detail will still be viewable. This ensures you get a comprehensive view of any report. This does not work both ways: when viewing a Workspace with only Models and no Reports, it will only show the Model detail since there are no Reports within that Workspace. If you do not want this perspective and prefer that Model detail only show in the Workspaces they are in, then set the All-Pages filter "Model in Workspace Flag" to TRUE.
- For backing up Reports & extracting the metadata, this mirrors what you can do at powerbi.com. This means that if you cannot download the report online, then the script will also not be able to download it. For Models, this works differently and if it's within a Premium, PPU, or Fabric capacity, even XMLA-only models can be backed up and extracted by leveraging the XMLA endpoint connection.

## Screenshots of Final Output
..
..

<img width="1235" alt="image" src="https://github.com/user-attachments/assets/805d3145-8290-4d84-8da2-bb27529bb050">
<img width="1259" alt="image" src="https://github.com/user-attachments/assets/54212360-8d0f-44c5-9337-db2cdd0fb5ee">
<img width="1240" alt="image" src="https://github.com/user-attachments/assets/488fc303-a9fa-4d4e-b0ce-c827fb440e83">
<img width="1259" alt="image" src="https://github.com/user-attachments/assets/9280e350-8714-40e5-8e09-d1de07faf5f5">
<img width="1221" alt="image" src="https://github.com/user-attachments/assets/e120c1bb-b52a-4197-aeb3-2a6ddbb67a9f">
<img width="1221" alt="image" src="https://github.com/user-attachments/assets/c9f5331d-8976-4f66-be76-5628e38e8d0f">
<img width="1241" alt="image" src="https://github.com/user-attachments/assets/9d814034-494d-478b-b231-f759d7eebeab">
