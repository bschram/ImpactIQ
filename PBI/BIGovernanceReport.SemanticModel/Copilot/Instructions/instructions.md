# Impact IQ — AI Instructions

## Purpose

This semantic model tracks **where every object from your Power BI semantic models is used** across the entire reporting environment. It answers questions like:

- "Where do I use [object name]?" → Which workspaces, reports, pages, and visuals reference it
- "What measures are affected if I change [object name]?" → Upstream (parent) and downstream (dependent) measure lineage
- "What objects are unused?" → Objects defined in models but never referenced in any report

## How to Interpret User Questions

When a user asks about a **measure**, **column**, **field**, **calculated column**, **hierarchy**, or **hierarchy level**, they are referring to the **value** in the `'Model Object Hierarchy'[Object Name]` column. The user provides the actual name (e.g. "Total Sales", "Customer ID", "FiscalYear") — **not** a type keyword.

### Filtering Rules

| User says... | Filter to apply |
|---|---|
| "Where do I use **Total Sales**?" | `'Model Object Hierarchy'[Object Name] = "Total Sales"` |
| "Where is the **Revenue** measure used?" | `'Model Object Hierarchy'[Object Name] = "Revenue"` AND `'Model Object Hierarchy'[Object Type] = "Measure"` |
| "Where is the **Customer ID** column used?" | `'Model Object Hierarchy'[Object Name] = "Customer ID"` AND `'Model Object Hierarchy'[Object Type] = "Column"` |
| "Where is the **FiscalYear** field used?" | `'Model Object Hierarchy'[Object Name] = "FiscalYear"` (field = any type) |
| "What depends on **Total Sales**?" | `'Model Object Hierarchy'[Object Name] = "Total Sales"` then read `Parent Measures - All` |
| "What feeds into **Total Sales**?" | `'Model Object Hierarchy'[Object Name] = "Total Sales"` then read `Dependent Measures - All` |

### How to Answer "Where is X Used?" (The Full Pattern)

Filtering `'Model Object Hierarchy'[Object Name]` alone only identifies the object. To get the **actual report locations** where that object is used, you must also evaluate measures that read from the `'All Reports'` fact table joined through `'Report Hierarchy'`:

1. **Filter** `'Model Object Hierarchy'[Object Name]` to the user's value (and optionally `[Object Type]`).
2. **Evaluate `[Report Objects Count]`** — this counts the usage instances across all reports for the filtered object.
3. **Group by `'Report Hierarchy'` columns** to break down WHERE it is used:
   - `'Report Hierarchy'[Workspace Name]` → which workspaces
   - `'Report Hierarchy'[Report Name]` → which reports
   - `'Report Hierarchy'[Page Name]` → which pages
   - `'All Reports'[Visual Name]` or `'All Reports'[Visual Type]` → which visuals
   - `'All Reports'[Filter Type Group]` → whether it's used as a visual field, page filter, report filter, or visual filter
4. **Use the "(Visual Impact)" measures** for summary counts:
   - `[Workspace Distinct Count (Visual Impact)]` → how many workspaces
   - `[Report Distinct Count (Visual Impact)]` → how many reports
   - `[Page Distinct Count (Visual Impact)]` → how many pages
   - `[Visual Distinct Count (Visual Impact)]` → how many visuals
   - `[Model Distinct Count (Visual Impact)]` → how many models

**Example DAX pattern** for "Where do I use Total Sales?":
```dax
EVALUATE
SUMMARIZECOLUMNS(
    'Report Hierarchy'[Workspace Name],
    'Report Hierarchy'[Report Name],
    'Report Hierarchy'[Page Name],
    KEEPFILTERS( 'Model Object Hierarchy'[Object Name] = "Total Sales" ),
    "Usage Count", [Report Objects Count]
)
```

### How to Answer "If I Change X, What Will Be Affected?" (Downstream / Consumers)

Use the `'Measure Lineage - Parents'` table. **Parents** are the measures that REFERENCE the selected measure in their own formulas — they are the **consumers/downstream** that would break if you change X.

1. **Filter** `'Model Object Hierarchy'[Object Name]` to the user's value.
2. **Group by** columns from `'Measure Lineage - Parents'` to show what would be affected.
3. **Level = Direct** means the parent measure directly references X in its DAX. **Indirect** means it depends on X through a chain of other measures.

**Example DAX pattern** for "If I change % CSAT 1 Star, what will be affected?":
```dax
EVALUATE
SUMMARIZECOLUMNS(
    'Measure Lineage - Parents'[Parent Measure Drillthrough Level Type],
    'Measure Lineage - Parents'[Parent Measure Level],
    'Measure Lineage - Parents'[Parent Measure Name],
    KEEPFILTERS( 'Model Object Hierarchy'[Object Name] = "% CSAT 1 Star" ),
    "Count", [Parent Measures Distinct Count - All]
)
ORDER BY
    'Measure Lineage - Parents'[Parent Measure Level] ASC
```

**CRITICAL:** Only list measures that appear in the query results. If the query returns 7 parent measures, report exactly 7 — do NOT add other measures from the model. The `'Measure Lineage - Parents'` table is the single source of truth for downstream impact.

### How to Answer "What Does X Rely On?" / "What Does X Depend On?" (Upstream / Inputs)

Use the `'Measure Lineage - Dependents'` table. **Dependents** are the measures that X references in its own formula — they are the **inputs/upstream** that X relies on.

**Example DAX pattern** for "What does % CSAT 1 Star rely on?":
```dax
EVALUATE
SUMMARIZECOLUMNS(
    'Measure Lineage - Dependents'[Dependent Measure Drillthrough Level Type],
    'Measure Lineage - Dependents'[Dependent Measure Level],
    'Measure Lineage - Dependents'[Dependent Measure Name],
    KEEPFILTERS( 'Model Object Hierarchy'[Object Name] = "% CSAT 1 Star" ),
    "Count", [Dependent Measures Distinct Count - All]
)
ORDER BY
    'Measure Lineage - Dependents'[Dependent Measure Level] ASC
```

**CRITICAL:** Only list measures that appear in the query results. Do NOT guess or infer additional measures.

### Key Principles

1. **Object Name is the lookup field.** The user's value goes into `'Model Object Hierarchy'[Object Name]`. Never search for the words "measure", "column", or "field" themselves — those are type qualifiers only.
2. **Object Type is the optional type filter.** If the user specifies "measure", "column", "hierarchy", etc., also filter `'Model Object Hierarchy'[Object Type]` to that value (Measure, Column, Hierarchy, Level, Relationship, Partition, CalculationGroup).
3. **"Field" means any type.** If the user says "field" without specifying measure/column, do NOT filter Object Type — it could be any object.
4. **Report locations come from Report Hierarchy + Report Objects Count.** You MUST use `[Report Objects Count]` (or other report measures) grouped by `'Report Hierarchy'` columns to show the actual workspace/report/page/visual where the object is used. Filtering `'Model Object Hierarchy'[Object Name]` alone does NOT give you report locations — it only identifies the object.
5. **Impact counts use the "(Visual Impact)" measures.** These count distinct workspaces, models, reports, pages, and visuals that reference the filtered object.
6. **Lineage direction matters.** Parents = downstream consumers (what breaks if I change this). Dependents = upstream inputs (what this measure relies on). **NEVER guess or infer lineage — only report what the query returns.**
7. **Used/Unused is on Model Object Hierarchy.** The `[Object Used Flag]` column shows "Used" or "Unused". The `[Object Used In]` column shows where (Reports, Measures, Reports & Measures).
8. **NEVER truncate URLs.** When returning URL links (Workspace Page URL, Workspace Report URL, etc.), always display the full, complete URL. Do not shorten, abbreviate, or use ellipsis in any URL.
9. **NEVER hallucinate data.** Only return results that come from executing a query against the model. If a lineage query returns 2 measures, report exactly 2 — do not add measures that "seem related" or that you expect to be there based on naming patterns.