from __future__ import annotations

from pathlib import Path
from typing import Any

from .models import LoadedWorkbook, SheetSpec


def inspect_sheet(path: Path, spec: SheetSpec) -> dict[str, Any]:
    return {
        "path": str(path),
        "sheet_id": spec.sheet_id,
        "title": spec.title,
        "freeze_rows": spec.freeze_rows,
        "freeze_cols": spec.freeze_cols,
        "zoom": spec.zoom,
        "column_widths": spec.column_widths,
        "row_heights": spec.row_heights,
        "cells": [cell.model_dump(mode="json", exclude_none=True) for cell in spec.cells],
        "ranges": [item.model_dump(mode="json", exclude_none=True) for item in spec.ranges],
        "merges": spec.merges,
        "tables": [item.model_dump(mode="json", exclude_none=True) for item in spec.tables],
        "charts": [item.model_dump(mode="json", exclude_none=True) for item in spec.charts],
        "assets": [cell.image_path for cell in spec.cells if cell.image_path],
    }


def inspect_workbook(loaded: LoadedWorkbook) -> dict[str, Any]:
    return {
        "project_root": str(loaded.project_root),
        "workbook": loaded.spec.model_dump(mode="json"),
        "theme": loaded.theme.model_dump(mode="json") if loaded.theme else None,
        "sheets": [inspect_sheet(sheet.path, sheet.spec) for sheet in loaded.sheets],
        "referenced_assets": sorted(
            {
                cell.image_path
                for sheet in loaded.sheets
                for cell in sheet.spec.cells
                if cell.image_path
            }
        ),
    }
