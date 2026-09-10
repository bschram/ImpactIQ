# Data coverage - every sheet, its source endpoint and the permission it needs

All four workbooks keep the v2 names, sheet names and columns; v3 adds sheets and columns (marked **new**) and never
removes any. Every sheet listed in `Config\SheetContract.json` is created even when empty, so the PBIT always loads.

Permission legend (workspace roles unless stated): **V** Viewer, **C** Contributor, **M** Member, **A** Admin;
"Build" = Build/Explore permission on the semantic model (Contributor+ have it implicitly). REST paths are relative to
`<ApiPrefix>/v1.0/myorg/` (Power BI) or `<FabricApiPrefix>/v1/` (Fabric). "My Workspace" rows use the same routes without
`groups/{id}/`. A 401/403/404 on an optional collector yields an empty sheet plus one Debug log line; core collectors
record the error in `InventoryErrors` and the workspace is re-collected on resume.

GCC note: Fabric REST is not offered in GCC (moderate). Sheets marked **Fabric** are empty there; everything else is
the same route on `api.powerbigov.us`.

## 1. `Power BI Environment Detail.xlsx` (stage `Inventory`, plus `Extras`)

### 1.1 Sheets from v2 (same order, same columns; additive columns noted)

| Sheet | Source endpoint | Permission | Notes / additive columns |
|---|---|---|---|
| `Workspaces` | `GET groups?$top=5000` (paged with `$skip`) | any role | + `WorkspaceCapacityId`, `CapacityName` (from Capacities), `WorkspaceState`, `WorkspaceHasWorkspaceLevelSettings`, `WorkspaceDefaultDatasetStorageFormat`, `WorkspaceIsSynthetic`, `WorkspaceApiScope` (pseudo rows "My Workspace", "Shared Reports (No Workspace Access)") |
| `FabricItems` **Fabric** | `GET workspaces/{id}/items` | V | + `FabricItemSensitivityLabelId`, `FabricItemFolderId`; Report/SemanticModel items excluded (their label ids are on the Reports/Datasets rows) |
| `Connections` **Fabric** | `GET connections` | connection User+ | + `ConnectionDetailsJson` |
| `Gateways` **Fabric** | `GET gateways` | gateway role | |
| `ItemConnections` **Fabric** | `GET workspaces/{id}/items/{itemId}/connections` | read+write on the item (C+) | blocked for protected sensitivity labels |
| `Datasets` | `GET groups/{id}/datasets` | V (Read-only callers get a trimmed object) | + `DatasetCapacityId`, `DatasetSensitivityLabelId`; nested API values (`upstreamDatasets`, ...) are flattened to JSON text |
| `DatasetSourcesInfo` | `GET groups/{id}/datasets/{dsId}/datasources` | V | |
| `DatasetRefreshHistory` | `GET groups/{id}/datasets/{dsId}/refreshes` | Learn: Write on the dataset (C+); Viewer often works | last 20-60 entries only (API retention); `refreshAttempts` flattened to JSON; attempted at Debug level for non-refreshable datasets |
| `DatasetRefreshSchedule` | `GET groups/{id}/datasets/{dsId}/refreshSchedule` | V | one row per day/time; + `DatasetRefreshScheduleKind` (`Import` / `Unavailable`) |
| `Dataflows` | `GET groups/{id}/dataflows` (+ **Fabric** `GET workspaces/{id}/dataflows` for Gen2 CI/CD) | V | + `DataflowSensitivityLabelId`, `DataflowFolderId` (Gen2) |
| `DataflowLineage` | `GET groups/{id}/datasets/upstreamDataflows` | V | + `DataflowWorkspaceId` (cross-workspace lineage) |
| `DataflowSourcesInfo` | `GET groups/{id}/dataflows/{dfId}/datasources` | V | skipped for Gen2 CI/CD (no objectId) |
| `DataflowRefreshHistory` | `GET groups/{id}/dataflows/{dfId}/transactions` | V | skipped for Gen2 CI/CD |
| `Reports` | `GET groups/{id}/reports` | V | + `ReportSensitivityLabelId`, `ReportHasSensitivityLabel`; pages not requested for paginated reports |
| `ReportPages` | `GET groups/{id}/reports/{rId}/pages` | V | My Workspace rows now carry `ReportId`/`ReportName` (v2 bug fixed) |
| `Apps` | `GET apps` | app access | apps whose workspace is in scope |
| `AppReports` | `GET apps/{appId}/reports` | app access | |

### 1.2 New sheets (all non-admin; v3)

| Sheet | Source endpoint | Permission | Columns |
|---|---|---|---|
| `Dashboards` | `GET groups/{id}/dashboards` | V | `DashboardId, DashboardName, DashboardIsReadOnly, DashboardWebUrl, DashboardEmbedUrl, WorkspaceId, WorkspaceName` |
| `DashboardTiles` | `GET groups/{id}/dashboards/{dId}/tiles` | V | `TileId, TileTitle, TileSubTitle, TileRowSpan, TileColSpan, TileEmbedUrl, ReportId, DatasetId, DashboardId, DashboardName, WorkspaceId, WorkspaceName` - dashboard -> report -> dataset lineage, dead tiles |
| `Capacities` | `GET capacities` | capacity admin or contributor (others get an empty list) | `CapacityId, CapacityDisplayName, CapacitySku, CapacityState, CapacityRegion, CapacityAdmins (";"-joined), CapacityUsersAccessRight` |
| `WorkspaceUsers` | `GET groups/{id}/users` | M+ (undocumented; 403 -> empty for that workspace) | `UserEmailAddress, UserDisplayName, UserIdentifier, UserPrincipalType, UserGroupUserAccessRight, UserGraphId, WorkspaceId, WorkspaceName` |
| `DatasetUsers` | `GET groups/{id}/datasets/{dsId}/users` | ReadWriteReshare on the dataset (owner / A / M / C with Reshare) | `UserIdentifier, UserPrincipalType, UserDatasetUserAccessRight, UserDisplayName, UserEmailAddress, DatasetId, DatasetName, WorkspaceId, WorkspaceName` - direct Build/Reshare grants |
| `DatasetParameters` | `GET groups/{id}/datasets/{dsId}/parameters` | V | `ParameterName, ParameterType, ParameterIsRequired, ParameterCurrentValue, ParameterSuggestedValues, DatasetId, DatasetName, WorkspaceId, WorkspaceName` - unsupported for XMLA-modified and some DirectQuery models (empty, not an error) |
| `DatasetDirectQueryRefreshSchedule` (Excel tab `DatasetDQRefreshSchedule`, 31-char limit) | `GET groups/{id}/datasets/{dsId}/directQueryRefreshSchedule` (only when the import schedule is unavailable) | V | `DQFrequency, DQLocalTimeZoneId, DQDay, DQTime, DatasetId, DatasetName, WorkspaceId, WorkspaceName` |
| `RunSummary` | `manifest.json` | - | `RunId, Stage, Status, StartedUtc, EndedUtc, DurationSeconds, ItemsDone, ItemsFailed, Error, Environment, Auth, RunMode, Machine, User, PSVersion, IsAzureDevOps, ResumeCount` - one row per stage + a `(Run)` row |
| `Failures` | `manifest.failures` | - | `Stage, ItemKey, Item, Message, TimeUtc, RunId` |
| `InventoryErrors` | collector errors recorded in `global.json` / `ws-*.json` | - | `WorkspaceId, WorkspaceName, Collector, Path, Message` - which optional collector returned what (403s, unsupported parameters, ...) |

### 1.3 Extras sheets (`-IncludeAdminApis`, Fabric administrator only)

The stage first probes `GET admin/capacities?$top=1`; a 401/403 means "not a Fabric admin" and all admin collectors are
skipped with one Warn. Raw responses are kept under `State\runs\<RunId>\extracts\admin\` and reused on resume.

| Sheet | Source endpoint | Notes |
|---|---|---|
| `AdminWorkspaces`, `AdminWorkspaceUsers` | `GET admin/groups?$top=5000&$expand=users,reports,datasets,dataflows,dashboards` (paged with `$skip`; 200 req/h) | every workspace in the tenant incl. ones the account is not in; users with roles |
| `ScanWorkspaces`, `ScanDatasets`, `ScanTables`, `ScanColumns`, `ScanMeasures`, `ScanExpressions`, `ScanDatasources`, `ScanReports`, `ScanDashboards`, `ScanDataflows`, `ScanUsers` | Scanner API: `GET admin/workspaces/modified` -> `POST admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True` in batches of 100 -> poll `admin/workspaces/scanStatus/{id}` -> `GET admin/workspaces/scanResult/{id}` | endorsement, sensitivity label, configuredBy, contentProviderType, model schema and DAX/M expressions (needs the tenant settings *Enhance admin APIs responses with detailed metadata / DAX and mashup expressions*), artifact users. Options: `AdminScanScopeOnly`, `AdminScanMaxWorkspaces`, `AdminScanBatchSize`, `AdminScanTimeoutMinutes` (via `$IQ.Options`) |
| `ActivityEvents` | `GET admin/activityevents?startDateTime=...&endDateTime=...` one call per UTC day for the last `-ActivityDays` (API window 28 days; `continuationUri` paging; 200 req/h) | one checkpoint per day (`admin-activity-<date>`); sheet capped at `ActivityMaxRows` (the per-day JSON files hold everything) |

### 1.4 Usage sheets (`-IncludeUsageMetrics`, non-admin)

| Sheet | Source | Permission | Notes |
|---|---|---|---|
| `UsageReportViews`, `UsageReportPageViews` | per workspace: find the hidden dataset `Report Usage Metrics Model` / `Usage Metrics Report` in `GET groups/{id}/datasets`, then `POST groups/{id}/datasets/{dsId}/executeQueries` with `INFO.TABLES()` / `INFO.COLUMNS()` discovery and `'Report views'`, `'Report page views'`, `'Reports'`, `'Users'` over the last `UsageDays` (30) | C+ in the workspace (creates/sees the usage model) with a Pro/PPU license; tenant settings *Usage metrics for content creators* and *Semantic Model Execute Queries REST API* on; a user must have opened **View usage metrics** once per workspace | rolling 30-day window; `Unnamed User [...]` when per-user data is disabled; not available for My Workspace; Microsoft's newer `executeDaxQueries` API excludes usage models - the tool uses the classic `executeQueries` |

## 2. `Report Detail.xlsx` (stages `ReportBackup` -> `ReportDetail` -> `Assemble`)

Rows are produced offline by the two Tabular Editor 2 C# scripts (`Report Detail Extract Script-PBIR.csx`, then
`Report Detail Extract Script.csx`) from every `.pbix` in `Report Backups\<RunId>\`; each script writes one tab-separated
`<Sheet>.txt`, and `Assemble` turns every `*.txt` into a sheet. Paginated reports (`.rdl`) are backed up only.

| Sheet | Content | How the PBIX was obtained | Permission |
|---|---|---|---|
| `Visuals`, `VisualObjects`, `VisualFilters`, `VisualInteractions`, `Pages`, `PageFilters`, `ReportFilters`, `Bookmarks`, `Connections`, `CustomVisuals`, `ReportLevelMeasures` | visual-level field usage, layouts/wireframe, filters, interactions, bookmarks, model connections, custom visuals, report-level measures (12-column layout; the v2 10-column bug is fixed) | 1. `GET groups/{id}/reports/{rId}/Export?downloadType=IncludeModel` (Pro) / `LiveConnect` (dedicated capacity); 2. fallback **Fabric** `POST workspaces/{id}/reports/{rId}/getDefinition` (PBIR / PBIR-Legacy, staged as a PBIX with a synthesized `Connections` file); 3. Pro `.bim` extracted from the IncludeModel PBIX with pbi-tools (needs Power BI Desktop on the machine) | Export: C+ in the report's workspace (and in the dataset's workspace for cross-workspace reports), tenant setting *Download reports* on; not for large-storage-format / Direct Lake / incremental-refresh / usage-metrics / template-app reports. getDefinition: read+write on the report, blocked for encrypted labels, Fabric only (not GCC) |
| `ReportExports` **new** | one row per report: `ReportName, ReportID, ModelID, WorkspaceID, WorkspaceName, ReportDisplayName, ReportType, FileName, ExportMethod (IncludeModel / LiveConnect / getDefinition), Status, Message, FileSizeBytes, DurationSec, DefinitionFormat, ModelExtract, BimPath, ReportDate` | rebuilt from the `ReportBackup` checkpoints (`Report Backups\<RunId>\ReportExports.txt`) | - |
| `ExtractErrors` (when present) | per-report extraction errors from the two csx scripts: `ReportName, Script (PBIR / Classic), Stage (unzip / report), Error, ReportDate` | written by the extractors when a report cannot be unzipped or parsed (the other reports are still processed and flushed to the TXT files one report at a time) | - |

## 3. `Model Detail.xlsx` (stages `ModelBackup` -> `ModelDetail` -> `Assemble`)

| Sheet | Content | Method | Permission / prerequisites |
|---|---|---|---|
| `Semantic Models` | one row per table, column, calculated column, measure, hierarchy, level, partition (M/DAX source), calculation group/item, RLS filter, relationship - 20 columns (`Type, Table, Name, FormatString, DisplayFolder, Description, IsHidden, TableStorageMode, Expression, ModelAsOfDate, ModelName, ModelID, Relationship*`) | **TabularEditor**: XMLA export of the model to `<Ws> ~ <Model>.bim` (`Provider=MSOLAP;Data Source=<XmlaPrefix>/v1.0/myorg/<workspace>;Password=<token>`), then `Model Detail Extract Script.csx`; Pro workspaces: `.bim` from the IncludeModel PBIX via pbi-tools. **Dax**: `POST groups/{id}/datasets/{dsId}/executeQueries` with `INFO.TABLES(), INFO.COLUMNS(), INFO.MEASURES(), INFO.RELATIONSHIPS(), INFO.PARTITIONS(), INFO.ROLES(), INFO.TABLEPERMISSIONS(), INFO.CALCULATIONGROUPS(), INFO.CALCULATIONITEMS(), INFO.HIERARCHIES(), INFO.LEVELS()` (one call each, raw results cached under `extracts\dax\`) | XMLA: dedicated capacity (Premium/PPU/Fabric), capacity setting *XMLA Endpoint* Read or Read Write, tenant setting *Allow XMLA endpoints and Analyze in Excel*, **Build** on the model (Contributor+ recommended - Build-only callers see masked metadata), Tabular Editor 2 runnable (Windows, .NET Framework 4.7.2+). DAX: tenant setting *Semantic Model Execute Queries REST API*, **Build** on the model, works on Pro; measure/RLS/partition expressions are blank without Write (Contributor+); raw `INFO.*` is documented as unsupported on `executeQueries` and some tenants reject it (HTTP 400, error 3239575574) - then the item is `Failed` with that message |
| `Measure Dependencies` | direct dependencies of every measure / calculated column / calculation item (`ObjectName, ObjectType, DependsOn, DependsOnType, ModelAsOfDate, ModelName, ModelID`) | TabularEditor: `Measure Dependency Extract Script.csx`; Dax: `INFO.CALCDEPENDENCY()` mapped to the same columns | as above; `INFO.CALCDEPENDENCY` requires **Write** on the model |

`-ModelDetailMethod Auto` uses Tabular Editor when a `.bim` exists, TE2 works and the `.bim` database name equals the
file name (XMLA / pbi-tools exports renamed by the tool), otherwise the built-in `.bim` (TMSL) parser (`Bim`: the same
20/7 columns, no external tool, no API call; measure dependencies come from parsing the DAX expressions - direct
references, approximate), otherwise DAX; `Both` runs TE2 and falls back to Bim, then DAX per model; `Bim` uses the
parser only; `Dax` never needs Tabular Editor (Linux/hosted agents without the .NET tools). When Tabular Editor cannot
export a dedicated-capacity model over XMLA and a Fabric token exists, `ModelBackup` tries
`POST workspaces/{id}/semanticModels/{dsId}/getDefinition?format=TMSL` (needs read+write on the model; unverified on
GCC and shared capacity) to obtain the `.bim`. The DAX path uses `INFO.VIEW.TABLES/COLUMNS/MEASURES/RELATIONSHIPS()`
first (documented for `executeQueries`), then the raw `INFO.*` functions best-effort; parts a tenant rejects are listed
as "unavailable via REST" in the checkpoint, and measure expressions are blank when the caller lacks Write.
`ModelName` is always `<CleanWs> ~ <CleanModel>` and `ModelID` the dataset GUID (dedicated) or the file name (Pro), so
the PBIT joins are unchanged.

## 4. `Dataflow Detail.xlsx` (stage `Dataflows` -> `Assemble`)

| Sheet | Content | Source | Permission |
|---|---|---|---|
| `Sheet1` | one row per query: `Dataflow ID, Dataflow Name, Query Name, Query, Report Date, Workspace Name - Dataflow Name` (+ the five empty DataTable artefact columns `RowError, RowState, Table, ItemArray, HasErrors` the PBIT expects, + additive `Load Enabled`, `Is Hidden`, `Query Group`, `Attributes`) | Gen1: `GET groups/{id}/dataflows/{dfId}` (model.json, saved raw as `<Ws> ~ <Dataflow>.txt`, `pbi:mashup.document` tokenised); Gen2 CI/CD **Fabric**: `POST workspaces/{id}/dataflows/{dfId}/getDefinition` (LRO; `mashup.pq` saved byte-exact as `.pq`, all parts under `<name>.definition\`) | Gen1: V (edit rights were needed in v2 only for the export UI); Gen2: read+write on the dataflow, Fabric only |

The v2 regex parser was replaced by a tokenizer that handles `#"quoted names"`, attributes in brackets, nested `let`,
trailing semicolons and `shared` inside strings.

## 5. What each optional flag adds (summary)

| Flag | Adds | Needs | Cost |
|---|---|---|---|
| (default) | sections 1.1, 1.2, 2, 3, 4 | workspace Viewer (metadata) / Contributor (backups, expressions) | 1 call per list + 1-3 calls per dataset/report/dataflow |
| `-IncludeMyWorkspace` | the account's own workspace (datasets, reports, pages, refreshes, schedules, dashboards, tiles, parameters, users) and reports shared without workspace access (pseudo workspace `Shared Reports (No Workspace Access)`, inventory only) | - | small |
| `-IncludeUsageMetrics` | `UsageReportViews`, `UsageReportPageViews` | Contributor+, usage model created once per workspace, two tenant settings | 4-6 `executeQueries` calls per workspace (120/min/user limit honoured) |
| `-IncludeAdminApis` | `AdminWorkspaces`, `AdminWorkspaceUsers`, `Scan*`, `ActivityEvents` | Fabric administrator role (probed; otherwise skipped) | Scanner: ~3 calls per 100 workspaces + polling; activity: 1+ call per day (200/h cap) |
| `-ActivityDays N` | how many UTC days of activity (max 28) | with `-IncludeAdminApis` | |
| `-ModelDetailMethod Dax` | model detail without Tabular Editor / XMLA | Build + executeQueries tenant setting | 12 calls per model |
| `-ModelDetailMethod Bim` | model detail from the `.bim` files alone (XMLA / pbi-tools / Fabric getDefinition exports), no Tabular Editor and no API call | a `.bim` per model from `ModelBackup` / `ReportBackup` | none |
| `-TimeBudgetMinutes N` | nothing extra - stops cleanly before a job cap (exit 3, `Paused`, resumed next run) | - | - |

Not collected because no non-admin API exists: report/dashboard/app user lists, subscriptions, endorsement and
sensitivity-label **names** (only ids), tenant-wide activity, unused artifacts. These come only with `-IncludeAdminApis`.
