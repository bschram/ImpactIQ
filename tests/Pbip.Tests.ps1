# Pbip.Tests.ps1 - the PBIP semantic model under PBI\ must tolerate the blank cells the workbooks contain.
# Export-Excel writes an empty string as an empty cell, which Power Query reads as null; every Text.* call on a
# column that can be blank (Expression of a table/column/relationship row, dataflow query names, datasource
# connection details, page names) therefore needs a null guard or the whole load fails with
# "We cannot convert the value null to type Text" (run 2026-09-16, All Models).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Tables = Join-Path (Get-IQTestRepoRoot) 'PBI/BIGovernanceReport.SemanticModel/definition/tables'
    function Get-TableSource { param([string]$Name) return (Get-Content -LiteralPath (Join-Path $script:Tables ($Name + '.tmdl')) -Raw -Encoding UTF8) }
}

Describe 'PBIP Power Query null guards (All Models load error, 2026-09-16)' {
    It 'All Models guards [Expression] before Text.PositionOf in both partition parsers' {
        $m = Get-TableSource 'All Models'
        ([regex]::Matches($m, [regex]::Escape('text = if [Expression] = null then "" else (try Text.From([Expression]) otherwise "")'))).Count | Should -Be 2
        $m | Should -Not -Match ([regex]::Escape('try Text.From([Expression]) otherwise [Expression]'))
    }
    It 'All Dataflows guards [Query Name] before Text.Contains' {
        $m = Get-TableSource 'All Dataflows'
        $m | Should -Match ([regex]::Escape('let queryName = if [Query Name] = null then "" else [Query Name] in not Text.Contains(queryName'))
        $m | Should -Not -Match ([regex]::Escape('Text.Contains([Query Name]'))
    }
    It 'All Dataflow Sources and All Model Sources guard the connection details before Text.Contains' {
        (Get-TableSource 'All Dataflow Sources') | Should -Not -Match ([regex]::Escape('Text.Contains([DataflowDatasourceConnectionDetails]'))
        (Get-TableSource 'All Dataflow Sources') | Should -Match ([regex]::Escape('let details = if [DataflowDatasourceConnectionDetails] = null then "" else [DataflowDatasourceConnectionDetails] in'))
        (Get-TableSource 'All Model Sources') | Should -Not -Match ([regex]::Escape('Text.Contains([DatasetDatasourceConnectionDetails]'))
        (Get-TableSource 'All Model Sources') | Should -Match ([regex]::Escape('let details = if [DatasetDatasourceConnectionDetails] = null then "" else [DatasetDatasourceConnectionDetails] in'))
    }
    It 'Report Hierarchy guards [Page Name] before Text.Contains' {
        (Get-TableSource 'Report Hierarchy') | Should -Match ([regex]::Escape('Text.Contains(if [Page Name] = null then "" else [Page Name], "isting")'))
    }
    It 'the placeholder row of an empty workbook never reaches a relationship key (Dataflow Hierarchy blank-key error, 2026-09-17)' {
        # Assemble writes one all-blank row when a sheet has no data (as the legacy script did); Power Query reads it as a
        # row of nulls, and a null key on the one side of a relationship fails the whole load.
        $expressions = Get-Content -LiteralPath (Join-Path (Get-IQTestRepoRoot) 'PBI/BIGovernanceReport.SemanticModel/definition/expressions.tmdl') -Raw -Encoding UTF8
        $expressions | Should -Match ([regex]::Escape('#"Removed Placeholder Rows" = Table.SelectRows(#"Changed Type1", each [Dataflow ID] <> null and [Dataflow ID] <> "")'))
        $expressions | Should -Match ([regex]::Escape('#"Removed Placeholder Rows" = Table.SelectRows(#"Renamed Columns1", each [Type] <> null and [Type] <> "")'))
        $expressions | Should -Match ([regex]::Escape('each ([ObjectType] = "Measure")')) -Because 'Base Measure Dependencies drops the placeholder row through its measure-only filter'
        (Get-TableSource 'Dataflow Hierarchy') | Should -Match ([regex]::Escape('[#"Workspace Name - Dataflow Name - Query Name"] <> null'))
        (Get-TableSource 'Report Hierarchy') | Should -Match ([regex]::Escape('[UniquePageID] <> null'))
        (Get-TableSource 'Measure Lineage') | Should -Match ([regex]::Escape('#"Removed Blank Keys" = Table.SelectRows(#"Removed Duplicates", each [#"WorkspaceName - ModelName - ObjectType - TableName - ObjectName"] <> null'))
    }
    It 'no table applies a Text.* function directly to a column without a preceding null guard' {
        # Text.From(null) returns null without an error, so it is not scanned. Guards accepted: "[col] <> null and",
        # "if [col] = null then """ or an earlier Table.ReplaceValue(..., null, ..., {"col"}) (All Reports: Page Name / Visual Name).
        $offenders = New-Object System.Collections.Generic.List[string]
        $textCall = 'Text\.(PositionOf|Contains|Middle|BetweenDelimiters|Start|End|Length|Upper|Lower|Trim|Split|Replace|Range|StartsWith|EndsWith)\(\s*\[(?<col>[^\]]+)\]'
        foreach ($f in Get-ChildItem -LiteralPath $script:Tables -Filter '*.tmdl' -File) {
            $src = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            foreach ($m in [regex]::Matches($src, $textCall)) {
                $col = [regex]::Escape($m.Groups['col'].Value)
                $before = $src.Substring([math]::Max(0, $m.Index - 1500), [math]::Min(1500, $m.Index))
                $guarded = ($before -match ('\[' + $col + '\]\s*<>\s*null')) -or
                           ($before -match ('\[' + $col + '\]\s*=\s*null\s*then\s*""')) -or
                           ($before -match ('Table\.ReplaceValue\([^\n]*null,[^\n]*\{[^\n]*"' + $col + '"'))
                if (-not $guarded) { $offenders.Add($f.Name + ': ' + $m.Value) }
            }
        }
        $offenders | Should -BeNullOrEmpty
    }
}
