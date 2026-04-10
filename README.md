# Herndon

Herndon is a local-first Python CLI for building Excel workbooks from structured JSON specs.

It is designed for deterministic workbook generation rather than interactive spreadsheet editing. The intended workflow is:

1. Create a Herndon project on disk.
2. Define workbook and sheet specs as JSON.
3. Mutate those specs through the CLI when convenient.
4. Validate the workbook before build.
5. Render a fresh `.xlsx` artifact under `.herndon/builds/`.

Herndon keeps the workbook source of truth in normal files, not inside Excel binaries.

See [SPEC.md](SPEC.md) for the original product specification. This README explains the current implementation in practical terms.

## What Herndon Does

Herndon currently supports:

- project initialization
- workbook creation
- sheet creation and ordering
- workbook metadata updates
- individual cell writes
- bulk range writes
- formulas as raw Excel formula strings
- merges
- freeze panes
- Excel tables
- chart rendering for `bar`, `column`, `line`, `pie`, and `scatter`
- stacked and percent-stacked bar/column charts
- chart legend position, axis titles, labels, and per-series colors
- named themes and named styles
- embedded PNG/JPEG images
- workbook validation
- JSON inspection output for agent use

## What It Does Not Do

Current limitations:

- no native histogram chart type
- no round-trip editing of arbitrary existing `.xlsx` files
- no pivot tables, VBA, conditional formatting, or data validation
- formula parsing is intentionally lightweight; formulas are treated mostly as opaque strings
- rendered workbook inspection is not the source of truth; JSON specs are

## Installation

Herndon is packaged as a Python project with a console entry point.

Editable install:

```bash
pip install -e .
```

After that, the CLI is available as:

```bash
herndon
```

If you do not install it, you can still run commands in-process from Python or use the package from the local `src/` directory.

## Tech Stack

The implementation currently uses:

- `typer` for the CLI
- `pydantic` for schema validation
- `openpyxl` for workbook rendering
- `Pillow` for asset inspection and image generation/loading

## Project Layout

Running `herndon init` creates a project with this layout:

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

A typical workbook lives under `workbooks/<workbook_id>/`:

```text
workbooks/
  expenses_2026/
    workbook.json
    sheets/
      001-overview.json
      002-stacked-mix.json
      003-trend.json
```

Generated output is written to the build path declared in `workbook.json`, usually something like:

```text
.herndon/builds/expenses_2026/expenses_2026.xlsx
```

## Core Files

### `workbook.json`

This file defines workbook metadata, theme selection, sheet ordering, and output path.

Example:

```json
{
  "version": 1,
  "workbook_id": "expenses_2026",
  "title": "Expenses 2026",
  "theme": "brand",
  "sheets": [
    "sheets/001-overview.json",
    "sheets/002-stacked-mix.json"
  ],
  "build": {
    "output": ".herndon/builds/expenses_2026/expenses_2026.xlsx"
  }
}
```

### Sheet JSON

Each sheet file contains layout, content, formatting references, and drawable elements.

Main sections:

- `cells`
- `ranges`
- `merges`
- `tables`
- `charts`
- `column_widths`
- `row_heights`
- `freeze_rows`
- `freeze_cols`

### Theme JSON

Themes live under `.herndon/themes/` and define named styles used by cells and ranges.

Herndon currently validates that referenced styles exist in the selected theme.

## Data Model

### Cells

Cells target a single address:

```json
{
  "cell": "B8",
  "formula": "=SUM(B3:B7)",
  "style": "currency_total"
}
```

Supported cell fields:

- `cell`
- `value`
- `formula`
- `style`
- `image_path`
- `w`
- `h`

If `image_path` is used, `w` and `h` are required and treated as inches.

### Ranges

Ranges write a 2D block of values starting at an anchor cell:

```json
{
  "anchor": "A2",
  "data": [
    ["Category", "Amount"],
    ["Rent", 2400],
    ["Travel", 900]
  ],
  "row_styles": {
    "0": "header"
  },
  "col_styles": {
    "1": "currency"
  }
}
```

Ranges are rendered before individual cells. If a cell appears in both a range and `cells`, the explicit cell entry wins.

### Tables

Excel tables are supported through `tables` entries:

```json
{
  "table_id": "expense_table",
  "name": "ExpenseData",
  "ref": "A2:B8",
  "header_row": true,
  "auto_filter": true,
  "style": "TableStyleMedium2"
}
```

### Charts

Supported chart types:

- `bar`
- `column`
- `line`
- `pie`
- `scatter`

Supported chart options in the current implementation:

- `title`
- `anchor`
- `w`
- `h`
- `series`
- `stacked`
- `percent_stacked`
- `show_legend`
- `legend_position`
- `show_data_labels`
- `show_percent_labels`
- `x_axis_title`
- `y_axis_title`
- `value_format`

Series support:

- `label`
- `values`
- `categories`
- `color`

Example:

```json
{
  "chart_id": "stacked_bar",
  "chart_type": "bar",
  "title": "Spend Mix by Month",
  "anchor": "F2",
  "w": 8,
  "h": 5,
  "percent_stacked": true,
  "legend_position": "b",
  "show_data_labels": true,
  "x_axis_title": "Share of Spend",
  "y_axis_title": "Month",
  "series": [
    {
      "label": "Rent",
      "values": "'Stacked Mix'!$B$2:$B$5",
      "categories": "'Stacked Mix'!$A$2:$A$5",
      "color": "#2563EB"
    }
  ]
}
```

Notes:

- `stacked` and `percent_stacked` are mutually exclusive.
- stacked behavior is currently implemented for bar/column charts only.
- series labels and colors are written into the generated chart XML.
- `openpyxl.load_workbook()` does not always round-trip those chart details back into high-level objects cleanly, so XML-level output can be more reliable than object readback for verification.

### Embedding Charts

Charts are embedded by adding objects to the sheet's `charts` array. They are not stored in `workbook.json`.

At render time, Herndon:

1. writes the sheet data
2. builds the chart from the referenced ranges
3. anchors the chart at the chart's `anchor` cell
4. sizes it using `w` and `h` in inches

Minimal example inside a sheet file:

```json
{
  "sheet_id": "overview",
  "title": "Overview",
  "freeze_rows": 0,
  "freeze_cols": 0,
  "zoom": 100,
  "column_widths": {},
  "row_heights": {},
  "cells": [],
  "ranges": [
    {
      "anchor": "A1",
      "data": [
        ["Category", "Amount"],
        ["Rent", 2400],
        ["Payroll", 8500],
        ["Travel", 900]
      ]
    }
  ],
  "merges": [],
  "tables": [],
  "charts": [
    {
      "chart_id": "expense_pie",
      "chart_type": "pie",
      "title": "Spend by Category",
      "anchor": "D2",
      "w": 7,
      "h": 5,
      "series": [
        {
          "label": "Amount",
          "values": "'Overview'!$B$2:$B$4",
          "categories": "'Overview'!$A$2:$A$4",
          "color": "#0F766E"
        }
      ],
      "show_legend": true,
      "legend_position": "r",
      "show_data_labels": true,
      "show_percent_labels": true
    }
  ]
}
```

Important details:

- `anchor` is the top-left placement cell for the chart object on the sheet.
- `values` and `categories` should reference cells that will exist after range/cell writes complete.
- sheet-qualified ranges should usually use the displayed sheet title, for example `'Overview'!$B$2:$B$4`.
- charts are rendered after tables and merges.
- if you move the source data, you must update the chart references yourself.

The same pattern applies to line, bar, column, and scatter charts. The only major difference is the `chart_type` and optional stacking or axis settings.

### Images

Images are embedded using cell-anchored `cells` entries with `image_path`.

Example:

```json
{
  "cell": "H1",
  "image_path": "assets/logo.png",
  "w": 1.6,
  "h": 0.55
}
```

Supported source formats:

- PNG
- JPEG

Relative paths are resolved from the project root.

## Render Order

Within each sheet, Herndon renders in this order:

1. column widths and row heights
2. range writes
3. individual cell writes
4. merges
5. tables
6. charts
7. freeze panes

This order is important because it controls overwrite and placement behavior.

## CLI Overview

Top-level commands:

```text
herndon init
herndon validate
herndon inspect
herndon render
herndon themes
herndon new
herndon workbooks
herndon sheets
herndon assets
```

Most read-oriented commands support `--format json`.

## Common Workflows

### 1. Initialize a Project

```bash
herndon init ./my-project
cd ./my-project
```

### 2. Create a Workbook and Sheet

```bash
herndon new workbook expenses_2026
herndon new sheet workbooks/expenses_2026/workbook.json overview
```

### 3. Set Workbook Metadata

```bash
herndon workbooks set-title workbooks/expenses_2026/workbook.json "Expenses 2026"
herndon workbooks set-theme workbooks/expenses_2026/workbook.json brand
herndon workbooks set-output workbooks/expenses_2026/workbook.json .herndon/builds/expenses_2026/expenses_2026.xlsx
```

### 4. Add Cell and Range Data

Set one cell:

```bash
herndon sheets set-cell workbooks/expenses_2026/workbook.json overview A1 --value "2026 Expenditures" --style title
```

Set a formula:

```bash
herndon sheets set-cell workbooks/expenses_2026/workbook.json overview B8 --formula "=SUM(B3:B7)" --style currency_total
```

Set a block of data:

```bash
herndon sheets set-range workbooks/expenses_2026/workbook.json overview A2 \
  --data-json '[["Category","Amount"],["Rent",2400],["Payroll",8500],["Travel",900]]' \
  --row-styles '{"0":"header"}' \
  --col-styles '{"1":"currency"}'
```

### 5. Add a Table

```bash
herndon sheets add-table workbooks/expenses_2026/workbook.json overview expense_table \
  --ref A2:B8 \
  --name ExpenseData
```

### 6. Add a Chart

Pie chart:

```bash
herndon sheets add-chart workbooks/expenses_2026/workbook.json overview expense_pie \
  --type pie \
  --anchor D2 \
  --title "Spend by Category" \
  --series-json '[{"label":"Amount","values":"'\''Overview'\''!$B$3:$B$7","categories":"'\''Overview'\''!$A$3:$A$7","color":"#0F766E"}]'
```

Line chart:

```bash
herndon sheets add-chart workbooks/expenses_2026/workbook.json trend line_trend \
  --type line \
  --anchor D2 \
  --title "Total Spend Trend" \
  --legend-position r \
  --y-axis-title "Spend" \
  --series-json '[{"label":"Total Spend","values":"'\''Trend'\''!$B$2:$B$6","categories":"'\''Trend'\''!$A$2:$A$6","color":"#7C3AED"}]'
```

Percent-stacked bar/column chart:

```bash
herndon sheets add-chart workbooks/expenses_2026/workbook.json stacked_mix stacked_bar \
  --type bar \
  --anchor F2 \
  --title "Spend Mix by Month" \
  --percent-stacked \
  --legend-position b \
  --show-data-labels \
  --x-axis-title "Share of Spend" \
  --y-axis-title "Month" \
  --series-json '[{"label":"Rent","values":"'\''Stacked Mix'\''!$B$2:$B$5","categories":"'\''Stacked Mix'\''!$A$2:$A$5","color":"#2563EB"}]'
```

### 7. Freeze Panes

```bash
herndon sheets freeze workbooks/expenses_2026/workbook.json overview --rows 2 --cols 0
```

### 8. Validate

```bash
herndon validate workbooks/expenses_2026/workbook.json
herndon validate workbooks/expenses_2026/workbook.json --format json
```

### 9. Inspect

Inspect emits normalized JSON that is useful for agents and tooling:

```bash
herndon inspect workbooks/expenses_2026/workbook.json --format json
```

You can also inspect an individual sheet file.

### 10. Render

```bash
herndon render workbooks/expenses_2026/workbook.json
herndon render workbooks/expenses_2026/workbook.json --format json
```

Successful render writes:

- the `.xlsx` file
- a build `manifest.json`

## Sheet Management Commands

Implemented sheet commands:

- `herndon sheets list <workbook.json>`
- `herndon sheets add <workbook.json> --sheet <sheet.json> [--after <sheet-id>]`
- `herndon sheets remove <workbook.json> <sheet-id> [--delete-files]`
- `herndon sheets rename <workbook.json> <sheet-id> <new-sheet-id>`
- `herndon sheets duplicate <workbook.json> <sheet-id> <new-sheet-id>`
- `herndon sheets move <workbook.json> <sheet-id> --after <sheet-id>`
- `herndon sheets set-cell ...`
- `herndon sheets set-range ...`
- `herndon sheets add-table ...`
- `herndon sheets add-chart ...`
- `herndon sheets update-chart ...`
- `herndon sheets remove-element ...`
- `herndon sheets set-merge ...`
- `herndon sheets clear-merge ...`
- `herndon sheets freeze ...`

## Assets

List project assets:

```bash
herndon assets list --format json
```

Inspect one asset:

```bash
herndon assets inspect assets/logo.png --format json
```

Example output:

```json
{
  "ok": true,
  "path": "/abs/path/assets/logo.png",
  "width": 180,
  "height": 60,
  "format": "PNG"
}
```

## Themes

List installed themes inside the project:

```bash
herndon themes --format json
```

## Validation

Herndon currently validates:

- workbook and sheet schema shape
- duplicate sheet IDs
- duplicate table names across the workbook
- missing theme references
- invalid cell and range addresses
- invalid row/column dimension references
- formulas that do not start with `=`
- unresolved sheet references in formulas as warnings
- chart references to unknown sheets as warnings
- merge overlap with table ranges
- unknown named styles
- missing image assets

Validation returns non-zero if any errors are present.

## JSON Output and Error Model

Read-oriented commands support `--format json` and are intended to be machine-friendly.

Typical validation output:

```json
{
  "ok": false,
  "issues": [
    {
      "severity": "error",
      "code": "unknown_style",
      "path": "/abs/path/workbooks/demo/sheets/001-summary.json",
      "field": "/cells/0/style",
      "message": "Style 'header2' is not defined in theme 'brand'"
    }
  ]
}
```

CLI error categories used by the implementation:

- `usage_error`
- `schema_error`
- `validation_error`
- `asset_error`
- `render_error`
- `io_error`

## Practical Demo Workbook

The repository includes a practical demo project at:

- [demo/expenditures_project](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project)

Key files:

- [workbook.json](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/workbooks/expenses_2026/workbook.json)
- [001-overview.json](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/workbooks/expenses_2026/sheets/001-overview.json)
- [002-stacked-mix.json](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/workbooks/expenses_2026/sheets/002-stacked-mix.json)
- [003-trend.json](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/workbooks/expenses_2026/sheets/003-trend.json)
- [logo.png](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/assets/logo.png)
- [expenses_2026.xlsx](/Users/johnhorton/tools/ep/herndon/demo/expenditures_project/.herndon/builds/expenses_2026/expenses_2026.xlsx)

The demo workbook currently exercises:

- a pie chart
- a percent-stacked chart
- a line chart
- an Excel table
- formulas
- named styles
- an embedded image

## Tests

The project includes CLI and render tests under `tests/`.

Run them with:

```bash
pytest -q
```

Current coverage includes:

- project initialization
- workbook/sheet creation
- workbook metadata updates
- inspection golden output
- validation behavior
- table rendering
- pie chart rendering
- stacked/percent-stacked and line chart rendering
- asset inspection
- embedded image rendering

## Internal Layout

Main source files:

- `src/herndon/cli.py`
- `src/herndon/models.py`
- `src/herndon/project.py`
- `src/herndon/validation.py`
- `src/herndon/inspection.py`
- `src/herndon/renderer.py`
- `src/herndon/errors.py`

Rough responsibility split:

- `models.py`: schema objects
- `project.py`: loading workbooks, sheets, and themes
- `validation.py`: workbook/sheet validation
- `inspection.py`: normalized inspection output
- `renderer.py`: `.xlsx` generation
- `cli.py`: user-facing command surface

## Notes on GUI Verification

The `.xlsx` output is standard and can be opened in Excel-compatible tools, but GUI behavior depends on the installed office suite.

For this repository, direct XML inspection of the rendered workbook is sometimes a better compatibility check than reading chart objects back through `openpyxl`, because some chart features do not round-trip perfectly through its read API.
