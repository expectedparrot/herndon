# Herndon v1 Specification

## 1. Purpose

Herndon is a local, agent-facing Python CLI for creating and updating Excel workbooks programmatically.

It is designed for workflows where an external agent or script is responsible for reasoning about workbook content, while Herndon provides durable, deterministic primitives for:

- creating `.xlsx` workbooks
- defining sheets in a structured intermediate format
- placing cells, ranges, tables, and charts
- applying styles and themes consistently
- supporting formulas as raw strings
- inspecting and validating workbook state

Herndon is not a spreadsheet editor. It is a build tool for workbooks.

For v1, Herndon targets `.xlsx` generation only.

## 2. Design Principles

1. Filesystem first

- Projects and source definitions live on disk.
- Derived artifacts can be rebuilt from source files.
- No database or background service is required.

2. Agent-facing interface

- Every read command must support JSON output.
- Mutating commands should accept structured files rather than only ad hoc flags.
- Errors must be machine-readable and actionable.

3. Deterministic rendering

- The same input should produce the same workbook structure whenever practical.
- Cell placement is always explicit — no hidden layout engine.
- Herndon should avoid hidden application state.

4. Excel-native output

- The primary artifact is a standard `.xlsx` file usable in Microsoft Excel, LibreOffice Calc, and Google Sheets import flows.
- Herndon should preserve compatibility over exotic rendering features.

5. Clear separation of content and style

- Sheet content is represented as structured data.
- Themes and named styles are represented separately.
- Agents can change cell values and formulas without rewriting low-level style logic.

6. Inspectable intermediate representation

- Herndon should expose a canonical workbook spec that can be linted, diffed, and regenerated.
- Users should not need to reverse-engineer `.xlsx` internals to understand generated output.

## 3. Goals

### 3.1 Functional goals

- Initialize a Herndon project.
- Create workbooks from a structured spec.
- Add, remove, reorder, and update sheets.
- Place cells with values, formulas, and styles.
- Write ranges of tabular data efficiently.
- Support named Excel Tables with headers and auto-filter.
- Support charts referencing cell ranges.
- Support themes and named styles.
- Export `.xlsx`.
- Validate a workbook spec before rendering.
- Inspect workbooks and sheet specs in JSON form.

### 3.2 Non-goals for v1

- Full fidelity round-trip editing of arbitrary existing `.xlsx` files.
- Pivot tables.
- Macros or VBA.
- Data validation dropdowns.
- Conditional formatting rules.
- Sparklines.
- Shared formulas or array formulas.
- Password protection.
- Browser-based editing.

## 4. Terminology

- Project: A directory containing Herndon source files and a `.herndon/` metadata directory.
- Workbook: A logical workbook with stable identity inside a project.
- Workbook spec: Canonical structured source describing sheets, theme references, and metadata.
- Sheet: One tab in the workbook.
- Cell: A single cell addressed by column letter and row number (e.g., `B4`).
- Range: A rectangular block of cells addressed as `A1:D10`.
- Table: An Excel Table (ListObject) defined over a named range with headers and optional auto-filter.
- Chart: A chart object placed on a sheet, referencing one or more data ranges.
- Style: A named set of formatting attributes (font, fill, border, number format, alignment).
- Theme: Shared workbook-level style defaults including named styles and a color palette.
- Asset: An external file referenced by the workbook, such as an image.
- Render: The act of compiling a workbook spec into an `.xlsx` artifact.

## 5. Project Layout

Running `herndon init` creates:

```text
.herndon/
  config.json
  themes/
  builds/
  cache/
  logs/
workbooks/
assets/
```

Recommended layout for one workbook:

```text
workbooks/
  q2_report/
    workbook.json
    sheets/
      001-summary.json
      002-revenue.json
      003-expenses.json
assets/
  logos/
  images/
.herndon/
  config.json
  themes/
    brand.json
  builds/
    q2_report/
      q2_report.xlsx
      manifest.json
  cache/
  logs/
```

## 6. Source of Truth

The source of truth is the workbook spec and referenced local assets.

For v1:

- `workbook.json` is canonical for workbook metadata and sheet ordering
- per-sheet JSON files are canonical for sheet content
- generated `.xlsx` files under `.herndon/builds/` are derived artifacts
- caches and manifests are disposable

## 7. Core Invariants

### 7.1 Identity invariants

- `workbook_id` is the stable identity of a workbook.
- `sheet_id` is the stable identity of a sheet within a workbook.
- Ordering is separate from identity.

### 7.2 Build invariants

- A successful render must be reproducible from source files and assets alone.
- Validation must run before final write of the output artifact.
- Derived build manifests may be deleted and regenerated.

### 7.3 Path invariants

- Asset references must be project-relative or absolute.
- Relative paths are resolved from the project root unless explicitly documented otherwise.
- Mutating commands should normalize stored paths.

### 7.4 Interface invariants

- Read commands support `--format json`.
- Errors from CLI commands should have stable machine-readable codes.
- Commands should exit non-zero on validation or render failure.

## 8. User Model

Herndon assumes two primary users:

- Human developers who define themes, named styles, and project conventions.
- Agents that generate or modify workbook specs through the CLI.

The agent is responsible for deciding what data and formulas belong in each cell.
Herndon is responsible for making that decision executable and consistent.

## 9. Workbook Model

### 9.1 Workbook metadata

Example `workbook.json`:

```json
{
  "version": 1,
  "workbook_id": "q2_report",
  "title": "Q2 Financial Report",
  "theme": "brand",
  "sheets": [
    "sheets/001-summary.json",
    "sheets/002-revenue.json",
    "sheets/003-expenses.json"
  ],
  "build": {
    "output": ".herndon/builds/q2_report/q2_report.xlsx"
  }
}
```

Fields:

- `version`: Herndon schema version
- `workbook_id`: stable logical id
- `title`: human-readable title
- `theme`: theme name or path
- `sheets`: ordered list of sheet spec file paths
- `build.output`: preferred output path

### 9.2 Sheet model

Example sheet file:

```json
{
  "sheet_id": "summary",
  "title": "Summary",
  "tab_color": "#B45309",
  "freeze_rows": 1,
  "freeze_cols": 0,
  "zoom": 100,
  "column_widths": {
    "A": 28,
    "B": 14,
    "C": 14
  },
  "row_heights": {
    "1": 22
  },
  "cells": [...],
  "ranges": [...],
  "merges": ["A1:C1"],
  "tables": [...],
  "charts": [...]
}
```

Fields:

- `sheet_id`: stable id, used as internal reference
- `title`: display name shown on the tab
- `tab_color`: optional hex color for the sheet tab
- `freeze_rows`: number of rows to freeze from the top
- `freeze_cols`: number of columns to freeze from the left
- `zoom`: display zoom percentage (default 100)
- `column_widths`: map of column letter to width in characters
- `row_heights`: map of row number (as string) to height in points
- `cells`: list of individual cell definitions
- `ranges`: list of range write operations for bulk data
- `merges`: list of cell merge ranges as strings
- `tables`: list of Excel Table definitions
- `charts`: list of chart definitions

### 9.3 Cell model

A cell defines the value or formula at a specific address, with optional style.

```json
{
  "cell": "B4",
  "value": 125000,
  "style": "currency"
}
```

```json
{
  "cell": "B10",
  "formula": "=SUM(B2:B9)",
  "style": "currency_total"
}
```

```json
{
  "cell": "A1",
  "value": "Revenue",
  "style": "header"
}
```

Fields:

- `cell`: required cell address (e.g., `B4`)
- `value`: a string, number, boolean, or null; mutually exclusive with `formula`
- `formula`: a raw formula string starting with `=`; mutually exclusive with `value`
- `style`: optional named style token

Value type rules:

- strings are written as text
- numbers are written as numeric cells
- booleans are written as boolean cells
- null clears the cell
- dates should be provided as ISO 8601 strings (`YYYY-MM-DD`) and rendered as Excel date values with an appropriate number format

Formula rules:

- formulas are written as-is into the cell's formula field
- Herndon does not evaluate or transform formulas
- Herndon validates that formula strings begin with `=`
- Herndon optionally checks that cell references within the formula point to addresses that exist within the workbook's sheet set
- the agent is responsible for formula correctness

### 9.4 Range model

A range write operation places a rectangular block of data starting at an anchor cell.

```json
{
  "anchor": "A1",
  "data": [
    ["Quarter", "Revenue", "Expenses", "Net"],
    ["Q1", 310000, 190000, 120000],
    ["Q2", 415000, 230000, 185000],
    ["Q3", 390000, 210000, 180000],
    ["Q4", 480000, 260000, 220000]
  ],
  "row_styles": {
    "0": "header"
  },
  "col_styles": {
    "1": "currency",
    "2": "currency",
    "3": "currency"
  }
}
```

Fields:

- `anchor`: top-left cell of the range
- `data`: 2D array of values (rows of columns); no formulas in range writes
- `row_styles`: map of zero-based row index to named style
- `col_styles`: map of zero-based column index to named style

Ranges are written before individual cell definitions. Cell definitions override range values at the same address.

### 9.5 Table model

An Excel Table wraps a range with a name, headers, and optional auto-filter.

```json
{
  "table_id": "revenue_table",
  "name": "RevenueData",
  "ref": "A1:D5",
  "header_row": true,
  "auto_filter": true,
  "style": "TableStyleMedium2"
}
```

Fields:

- `table_id`: stable id within the sheet
- `name`: Excel table name (must be unique within the workbook, no spaces)
- `ref`: cell range string the table occupies
- `header_row`: whether the first row is a header (default true)
- `auto_filter`: whether to enable auto-filter (default true)
- `style`: Excel built-in table style name (e.g., `TableStyleMedium2`)

### 9.6 Chart model

A chart is placed on a sheet and references data ranges for its series.

```json
{
  "chart_id": "revenue_chart",
  "chart_type": "column",
  "title": "Quarterly Revenue",
  "anchor": "F2",
  "w": 8,
  "h": 5,
  "series": [
    {
      "label": "Revenue",
      "values": "Sheet1!B2:B5",
      "categories": "Sheet1!A2:A5"
    }
  ],
  "show_legend": true,
  "value_format": "#,##0"
}
```

Fields:

- `chart_id`: stable id within the sheet
- `chart_type`: `bar`, `column`, `line`, `pie`, `scatter`
- `title`: optional chart title
- `anchor`: top-left cell for chart placement
- `w`: chart width in inches
- `h`: chart height in inches
- `series`: list of data series definitions
- `show_legend`: whether to show a legend (default true)
- `value_format`: optional number format string for axis values

Series fields:

- `label`: series name shown in legend
- `values`: cell range reference for data values (sheet-qualified, e.g., `Sheet1!B2:B5`)
- `categories`: cell range reference for category labels

## 10. Style and Theme Model

Themes live under `.herndon/themes/`.

Example theme:

```json
{
  "name": "brand",
  "colors": {
    "background": "#FFFFFF",
    "text": "#1E1E1E",
    "accent": "#B45309",
    "muted": "#6B7280",
    "header_fill": "#F3F4F6"
  },
  "fonts": {
    "default": "Aptos",
    "heading": "Aptos Display"
  },
  "styles": {
    "header": {
      "font_name": "Aptos Display",
      "font_size": 11,
      "bold": true,
      "color": "#1E1E1E",
      "fill": "#F3F4F6",
      "border_bottom": "thin",
      "alignment": "left"
    },
    "body": {
      "font_name": "Aptos",
      "font_size": 11,
      "color": "#1E1E1E"
    },
    "currency": {
      "font_name": "Aptos",
      "font_size": 11,
      "number_format": "#,##0.00"
    },
    "currency_total": {
      "font_name": "Aptos",
      "font_size": 11,
      "bold": true,
      "number_format": "#,##0.00",
      "border_top": "thin"
    },
    "percentage": {
      "font_name": "Aptos",
      "font_size": 11,
      "number_format": "0.0%"
    },
    "date": {
      "font_name": "Aptos",
      "font_size": 11,
      "number_format": "YYYY-MM-DD"
    },
    "integer": {
      "font_name": "Aptos",
      "font_size": 11,
      "number_format": "#,##0"
    }
  }
}
```

### Style fields

Each named style may specify:

- `font_name`: font family
- `font_size`: point size
- `bold`: boolean
- `italic`: boolean
- `underline`: boolean
- `color`: hex font color
- `fill`: hex background fill color
- `number_format`: Excel number format string
- `alignment`: `left`, `center`, `right`, `general`
- `vertical_alignment`: `top`, `middle`, `bottom`
- `wrap_text`: boolean
- `border_top`, `border_bottom`, `border_left`, `border_right`: border style string (`thin`, `medium`, `thick`, `dashed`, `dotted`)
- `border_color`: hex border color (applies to all border sides unless overridden)

Styles defined in cells and ranges reference names from the active theme. If a style name is not found in the theme, Herndon raises a validation error.

## 11. CLI Surface

Primary commands:

```text
herndon init [path]
herndon new workbook <name>
herndon new sheet <workbook.json> <sheet-id>
herndon validate <workbook-or-sheet-path>
herndon render <workbook-path>
herndon inspect <workbook-or-sheet-path>
```

Sheet management:

```text
herndon sheets list <workbook.json>
herndon sheets add <workbook.json> --sheet <sheet.json> [--after <sheet-id>]
herndon sheets remove <workbook.json> <sheet-id> [--delete-files]
herndon sheets rename <workbook.json> <sheet-id> <new-sheet-id>
herndon sheets duplicate <workbook.json> <sheet-id> <new-sheet-id>
herndon sheets move <workbook.json> <sheet-id> --after <sheet-id>
```

Cell and range mutation:

```text
herndon sheets set-cell <workbook.json> <sheet-id> <cell> --value <value> [--style <style>]
herndon sheets set-cell <workbook.json> <sheet-id> <cell> --formula <formula> [--style <style>]
herndon sheets set-range <workbook.json> <sheet-id> <anchor> --data-json <json> [--row-styles <json>] [--col-styles <json>]
herndon sheets add-table <workbook.json> <sheet-id> <table-id> --ref <range> --name <name>
herndon sheets add-chart <workbook.json> <sheet-id> <chart-id> --type <chart_type> --anchor <cell> --series-json <json>
herndon sheets update-chart <workbook.json> <sheet-id> <chart-id> [--title <title>] [--show-legend]
herndon sheets remove-element <workbook.json> <sheet-id> <element-id>
herndon sheets set-merge <workbook.json> <sheet-id> <range>
herndon sheets clear-merge <workbook.json> <sheet-id> <range>
herndon sheets freeze <workbook.json> <sheet-id> [--rows <n>] [--cols <n>]
```

Resource discovery:

```text
herndon themes [--project-root <path>]
herndon assets list [--project-root <path>]
herndon assets inspect <asset-path>
```

Interface expectations:

- commands default to human-readable output
- `--format json` produces stable structured output for agents
- mutating commands should support `--dry-run` when practical
- `--project-root` overrides automatic project root discovery

## 12. Validation

`herndon validate` should check:

- schema validity of workbook and sheet specs
- duplicate `sheet_id` violations
- duplicate table names within the workbook
- missing theme references
- missing asset references
- cell address format validity
- range address format validity
- formula strings that do not start with `=`
- formula cell references that point to sheets not present in the workbook (warning, not error)
- out-of-range row/column references for explicit column widths and row heights
- merge ranges that overlap with Table ranges
- chart series ranges that reference unknown sheet names
- style names not defined in the active theme

Validation output should include:

- severity: `error` or `warning`
- code: stable string identifier
- message
- file path
- field path when available

Example:

```json
{
  "ok": false,
  "issues": [
    {
      "severity": "error",
      "code": "unknown_style",
      "path": "workbooks/q2_report/sheets/001-summary.json",
      "field": "/cells/3/style",
      "message": "Style 'highlight' is not defined in theme 'brand'"
    },
    {
      "severity": "warning",
      "code": "unresolved_sheet_ref",
      "path": "workbooks/q2_report/sheets/001-summary.json",
      "field": "/cells/5/formula",
      "message": "Formula references sheet 'Detail' which is not in this workbook"
    }
  ]
}
```

## 13. Render Semantics

`herndon render` behavior:

1. Load project config.
2. Load workbook spec and referenced sheet specs.
3. Resolve theme and named style references.
4. Validate the full workbook.
5. Materialize a new `.xlsx` in a temporary path.
6. Write a build manifest.
7. Atomically move the final artifact into place.

Render should fail without replacing the previous build artifact if validation or write fails.

### Render order within a sheet

For each sheet, the renderer applies in this order:

1. Column widths and row heights
2. Range writes (top to bottom, left to right)
3. Individual cell writes (overriding range values at the same address)
4. Merges
5. Tables
6. Charts
7. Freeze panes

This ordering is deterministic and documented so agents can reason about precedence.

## 14. Build Manifest

Each successful render should write a manifest such as:

```json
{
  "workbook_id": "q2_report",
  "source_path": "workbooks/q2_report/workbook.json",
  "output_path": ".herndon/builds/q2_report/q2_report.xlsx",
  "rendered_at": "2026-04-09T17:00:00Z",
  "sheet_count": 3,
  "theme": "brand"
}
```

The manifest is derived and disposable.

## 15. Inspection Commands

`herndon inspect` should expose normalized JSON for:

- project config
- workbook metadata
- sheet list and order
- cell inventory per sheet (addresses, values, formulas, styles)
- table definitions
- chart definitions
- referenced assets

The normalized inspection form is the main machine interface for agents that need to reason about current workbook state.

## 16. Asset Handling

Supported v1 assets:

- PNG
- JPEG

Asset rules:

- missing assets are validation errors
- asset dimensions should be inspectable through CLI
- Herndon should not mutate source assets

Image placement on sheets is supported as a cell-anchored element:

```json
{
  "cell": "E2",
  "image_path": "assets/logos/brand.png",
  "w": 2.0,
  "h": 0.8
}
```

Image placement fields:

- `cell`: anchor cell (top-left corner of the image)
- `image_path`: project-relative or absolute path
- `w`: width in inches
- `h`: height in inches

## 17. Python Implementation Constraints

The implementation should be a Python package with a console entry point.

Recommended v1 stack:

- CLI: `typer` or `click`
- schema validation: `pydantic`
- `.xlsx` generation: `openpyxl`
- image probing: `Pillow`

Design constraint:

- Herndon's internal spec should not mirror `openpyxl` too closely
- backend-specific details should be isolated behind a renderer layer

This allows future replacement or augmentation of the rendering backend without rewriting the CLI contract.

## 18. Error Model

CLI failures should map to stable categories:

- `usage_error`
- `schema_error`
- `validation_error`
- `asset_error`
- `render_error`
- `io_error`

JSON error shape:

```json
{
  "ok": false,
  "error": {
    "code": "validation_error",
    "message": "Workbook contains 2 validation errors",
    "details": []
  }
}
```

## 19. Testing Requirements

v1 should include:

- schema validation tests
- CLI command tests
- fixture-based render tests
- regression tests for cell geometry and formula pass-through
- golden-file tests for normalized inspection output

Excel binary diffs are noisy, so tests should prefer:

- normalized manifest assertions
- XML-part assertions for selected sheet content
- sheet count and cell value checks

## 20. Security and Safety

Herndon is a local tool, but it still must:

- avoid executing code from workbook specs
- treat formula strings as opaque data — write them to cells, do not evaluate them
- treat file references as data only
- reject path traversal where project-local paths are required
- avoid shelling out implicitly

## 21. Open Questions

The following are intentionally unresolved and should be decided before implementation:

1. Should the canonical sheet source format be JSON only, or JSON plus YAML?
2. Should Herndon support conditional formatting in v2, and what is the right spec shape for it?
3. Should chart series ranges be validated strictly (error) or loosely (warning) when they reference external sheet names?
4. Should `herndon render` support partial builds of a sheet subset for fast iteration?
5. Should named styles support inheritance (e.g., `currency_total` extends `currency`)?
6. Should image elements be a first-class sheet element type or remain a secondary concern in v1?

## 22. Recommended v1 Scope Cut

To keep v1 tractable, Herndon should ship with:

- project init
- workbook creation
- JSON sheet specs
- cells with values and raw formula strings
- range writes for bulk tabular data
- merges
- named Excel Tables
- bar, column, line, and pie charts
- freeze panes
- column widths and row heights
- theme files with named styles
- image placement
- validation
- JSON inspection
- deterministic `.xlsx` rendering

The first version should not attempt round-trip editing of existing `.xlsx` files. The agent-facing value comes from reliable generation, validation, and inspection.
