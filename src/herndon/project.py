from __future__ import annotations

from pathlib import Path

from .errors import HerndonError
from .models import LoadedSheet, LoadedWorkbook, ProjectConfig, SheetSpec, ThemeSpec, WorkbookSpec
from .utils import find_project_root, load_json


def require_project_root(start: Path, explicit_root: Path | None = None) -> Path:
    root = explicit_root.resolve() if explicit_root else find_project_root(start)
    if root is None:
        raise HerndonError("io_error", "Could not locate Herndon project root")
    return root


def load_project_config(project_root: Path) -> ProjectConfig:
    config_path = project_root / ".herndon" / "config.json"
    return ProjectConfig.model_validate(load_json(config_path))


def load_theme(project_root: Path, theme_ref: str | None) -> ThemeSpec | None:
    if not theme_ref:
        return None
    theme_path = Path(theme_ref)
    if not theme_path.suffix:
        theme_path = project_root / ".herndon" / "themes" / f"{theme_ref}.json"
    elif not theme_path.is_absolute():
        theme_path = project_root / theme_path
    if not theme_path.exists():
        return None
    return ThemeSpec.model_validate(load_json(theme_path))


def load_sheet(path: Path) -> LoadedSheet:
    return LoadedSheet(path=path, spec=SheetSpec.model_validate(load_json(path)))


def load_workbook(workbook_path: Path, project_root: Path | None = None) -> LoadedWorkbook:
    workbook_path = workbook_path.resolve()
    project_root = require_project_root(workbook_path.parent, explicit_root=project_root)
    workbook = WorkbookSpec.model_validate(load_json(workbook_path))
    sheets: list[LoadedSheet] = []
    for relative_path in workbook.sheets:
        sheet_path = (workbook_path.parent / relative_path).resolve()
        sheets.append(load_sheet(sheet_path))
    return LoadedWorkbook(
        project_root=project_root,
        path=workbook_path,
        spec=workbook,
        sheets=sheets,
        theme=load_theme(project_root, workbook.theme),
    )
