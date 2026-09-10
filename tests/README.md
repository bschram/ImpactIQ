# ImpactIQ tests

Pester 5 tests for the modular, headless ImpactIQ (`ImpactIQ.ps1` + `Config/Modules/ImpactIQ.*.ps1`). Everything runs
offline: the Power BI / Fabric / Entra endpoints are answered from fixtures, sleeps are mocked, and the Windows-only
paths (Tabular Editor, pbi-tools, DPAPI, WinForms) are skipped when not on Windows.

## Prerequisites

| Requirement | Windows (Windows PowerShell 5.1 or PowerShell 7) | Linux / macOS (PowerShell 7) |
|---|---|---|
| Pester 5.7.1 | `Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser -Force -SkipPublisherCheck` | same |
| PSScriptAnalyzer | `Install-Module PSScriptAnalyzer -Scope CurrentUser -Force` | same |
| ImportExcel (Assemble / Entry tests) | `Install-Module ImportExcel -Scope CurrentUser -Force` | same (`Get-ExcelSheetInfo`/`Import-Excel` work without libgdiplus; the AutoSize warning is harmless) |

The Windows PowerShell 5.1 that ships with Windows carries Pester 3.4 in `C:\Program Files\WindowsPowerShell\Modules`;
the `-SkipPublisherCheck` flag above lets the 5.7.1 install proceed. `tests/Invoke-Tests.ps1` imports Pester with
`-RequiredVersion 5.7.1`, so the old module never interferes.

## Running

The full quality gate (parse every `.ps1`, PSScriptAnalyzer with `PSScriptAnalyzerSettings.psd1`, Pester):

```powershell
# Windows PowerShell 5.1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Invoke-Tests.ps1

# PowerShell 7 (Windows, Linux, macOS)
pwsh -NoProfile -File ./tests/Invoke-Tests.ps1
```

Exit code `0` = everything passed, `1` = a syntax error, a blocking analyzer finding (Error severity, or a
`PSUseCompatibleSyntax` / `PSUseCompatibleCommands` warning) or a failed test. Results are also written as NUnit XML to
`tests/TestResults/ImpactIQ.Tests.xml` (publishable with the Azure DevOps `PublishTestResults@2` task, format `NUnit`).

Options:

```powershell
./tests/Invoke-Tests.ps1 -TestName Common,State      # only these test files
./tests/Invoke-Tests.ps1 -SkipAnalyzer               # parse + Pester only
./tests/Invoke-Tests.ps1 -SkipPester                 # parse + analyzer only
./tests/Invoke-Tests.ps1 -ResultsPath C:\temp\results
```

Running a single file straight through Pester (handy while developing):

```powershell
Import-Module Pester -RequiredVersion 5.7.1
Invoke-Pester -Path ./tests/Http.Tests.ps1 -Output Detailed
```

## Layout

| File | Covers |
|---|---|
| `TestHelpers.ps1` | Dot-sourced by every test file's `BeforeAll`: loads the modules into the test's script scope (so `$script:IQ` and `Mock` work), creates a temp BaseFolder with `Config\` placeholders (csx names, `Blank Model.bim`, `SheetContract.json`), the fixture-backed `Invoke-IQApi` stand-in (`Invoke-IQTestApiFixture`), JWT / HTTP response / error-record builders. |
| `Common.Tests.ps1` | `Get-IQCleanName` parity with the monolith regex, `Get-IQSafeKey`, JSON round trip, `Invoke-IQWithRetry`, endpoint table (section 3), logging/redaction, exit codes. |
| `Auth.Tests.ps1` | JWT expiry, AES token cache (DPAPI on Windows only), Auto-mode table, device-code state machine (`authorization_pending` -> `slow_down` -> success) with a mocked `Invoke-RestMethod`, ROPC AADSTS mapping, AccessToken mode, headless guards. |
| `Http.Tests.ps1` | URL building, `@odata.nextLink` / `continuationUri` / `continuationToken` paging, 429 Retry-After, 401 refresh-once, 403/404/400 -> `$null`, 5xx backoff, `Get-IQHttpErrorInfo` (WebException + HttpResponseException + 5.1-style response object), Fabric LRO polling, downloads. |
| `State.Tests.ps1` | Manifest schema, every resume branch of section 5.2 (with an injected clock), item checkpoints incl. missing-output invalidation, stage runner statuses, run summary. |
| `Inventory.Tests.ps1` | The Inventory stage against the API fixtures: `ws-*.json` shapes (rename maps, new collectors), `global.json`, My Workspace (C6-01), scope resolution for Workspaces / Reports (remote model workspace) / Models modes, headless guards, resume, failure handling. |
| `Dataflows.Tests.ps1` | Gen1 mashup / Gen2 `.pq` parsing (names with spaces, `#"..."`, attribute records, nested `let`, trailing `;`, `shared` inside strings) and the Dataflows stage (backups, extracts, checkpoints, resume). |
| `Dax.Tests.ps1` | `ConvertFrom-IQDaxRows`, `Invoke-IQDaxQuery` request/response handling, `Get-IQModelDetailViaDax` producing the exact csx CSV headers/rows from INFO.VIEW.* + INFO.* fixtures, the raw-INFO-rejected path, the `.bim` path. |
| `Bim.Tests.ps1` | TMSL parser and CSV export from `fixtures/bim/sample-model.bim`. |
| `Assemble.Tests.ps1` | The four workbooks built from a prepared run state; every sheet/column of `Config/SheetContract.json` is asserted with `Import-Excel`; TXT/CSV/dataflow-extract parsing; empty-state run. |
| `Entry.Tests.ps1` | `ImpactIQ.ps1`: parse + parameter set, `-NonInteractive` without scope fails with "no scope", `-Stages Assemble` on a prepared state folder (child process with network cmdlets shadowed). Skipped until `ImpactIQ.ps1` exists. |

### Fixtures (`tests/fixtures`)

* `api/<Api>__<path with / as __>.json` - Power BI / Fabric REST responses (groups, datasets, datasources, refreshes,
  refreshSchedule, directQueryRefreshSchedule, parameters, users, reports, pages, dataflows, upstreamDataflows,
  transactions, dashboards, tiles, apps, capacities, My Workspace `datasets`/`reports`, Fabric items/connections/
  gateways/dataflows). A file containing `{ "__status": 403 }` makes the stand-in return `$null`, like `Invoke-IQApi`
  does for 403/404/400. `ids.json` lists the GUIDs used across the fixtures.
* `dax/*.json` - `executeQueries` responses keyed by the DAX text (`EVALUATE INFO.VIEW.TABLES()` -> `info-view-tables.json`),
  raw `INFO.*` with integer enum codes, `INFO.CALCDEPENDENCY`, an engine error result.
* `dataflows/` - a Gen1 model.json with `pbi:mashup.document`, a Gen2 `mashup.pq` and `queryMetadata.json`.
* `extracts/report-detail/*.txt` - tab-separated files with the exact csx/contract columns (Pages, Visuals,
  ReportLevelMeasures, Connections, a header-only Bookmarks).
* `extracts/model-detail/*.csv` - a Semantic Models CSV and its `_MD.csv` with the exact csx headers.
* `bim/sample-model.bim` - a small TMSL model (tables, calculated column, measures, relationships, roles, calc group).

## Notes

* Pester evaluates `-Skip:` / `-TestCases` during discovery, before `BeforeAll`; the test files therefore compute
  those values at the top of the file and again in `BeforeAll` for the run phase.
* Header-only worksheets come back as zero rows from `Import-Excel`; use `Get-IQTestSheetHeader` (reads row 1 with
  `-NoHeader`) to assert their columns.
* Temporary folders live under `<temp>/ImpactIQ-tests/` and are removed by each file's `AfterAll`.
* The Entry tests run `ImpactIQ.ps1` in a child PowerShell (the entry point uses `exit`); `Invoke-WebRequest` and
  `Invoke-RestMethod` are shadowed by global functions there so no real HTTP call can happen, and the auth mode is
  `AccessToken` with a fake JWT from `IMPACTIQ_PBI_TOKEN`.
