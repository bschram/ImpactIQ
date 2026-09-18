# Copilot instructions for this repo

## Big picture
- This is a Power BI/Fabric governance and impact-analysis toolkit driven by a single PowerShell orchestrator: [Final PS Script.txt](Final%20PS%20Script.txt). It downloads tools, authenticates, scans workspaces, backs up artifacts, and produces Excel/CSV extracts used by the report template.
- C# Tabular Editor scripts in Config handle model/report metadata extraction. Examples: [Config/Model Detail Extract Script.csx](Config/Model%20Detail%20Extract%20Script.csx) writes model metadata CSVs; [Config/Report Detail Extract Script.csx](Config/Report%20Detail%20Extract%20Script.csx) parses report Layout/Connections JSON inside unpacked report folders.
- Output is consumed by the Power BI template [Power BI Governance Model.pbit](Power%20BI%20Governance%20Model.pbit).

## Critical workflows
- Primary run path: execute the PowerShell script in [Final PS Script.txt](Final%20PS%20Script.txt) (often renamed to .ps1). It auto-downloads pbi-tools and Tabular Editor 2 into Config and installs user-scoped PowerShell modules.
- The script prompts for cloud environment (Public/Germany/USGov/China/USGovHigh/USGovMil) and sets environment-specific API endpoints.
- Backups and extracts are written under dated folders (yyyy-MM-dd) in Model/Report/Dataflow backup roots; final Excel outputs are in repo root (e.g., Model Detail.xlsx, Report Detail.xlsx).

## Project conventions & patterns
- Base path is assumed to be C:\Power BI Backups. Many scripts use Directory.GetCurrentDirectory() as the base; keep that assumption if adding new extract steps.
- Tabular Editor scripts should write into the latest-dated folder under the appropriate backup root. Example in [Config/Model Detail Extract Script.csx](Config/Model%20Detail%20Extract%20Script.csx): Model Backups/yyyy-MM-dd/<ModelName>.csv.
- Report extraction expects unpacked PBIX/PBIR folders with Report\Layout and Connections files; see [Config/Report Detail Extract Script.csx](Config/Report%20Detail%20Extract%20Script.csx).

## Integrations & dependencies
- Power BI REST API via MicrosoftPowerBIMgmt PowerShell module (installed at run time in [Final PS Script.txt](Final%20PS%20Script.txt)).
- XMLA used for model backups where available; pbi-tools is used for Pro/non-XMLA scenarios and is downloaded into Config/PBI Tools.
- Tabular Editor 2 portable is downloaded into Config/TabularEditor and invoked from the PowerShell script.

## Where to look first
- Orchestration and workflow: [Final PS Script.txt](Final%20PS%20Script.txt)
- Model/Report extraction logic: [Config/Model Detail Extract Script.csx](Config/Model%20Detail%20Extract%20Script.csx), [Config/Report Detail Extract Script.csx](Config/Report%20Detail%20Extract%20Script.csx)
- Usage and setup: [README.md](README.md)
