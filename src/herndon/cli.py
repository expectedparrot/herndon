from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Optional

import typer
from PIL import Image
from pydantic import ValidationError

from .errors import HerndonError
from .inspection import inspect_sheet, inspect_workbook
from .models import ChartSpec, CellSpec, ProjectConfig, RangeSpec, SheetSpec, TableSpec, WorkbookSpec
from .project import load_sheet, load_workbook, require_project_root
from .renderer import render_workbook
from .utils import atomic_dump_json, dump_json, load_json, parse_cli_value
from .validation import validate_path

app = typer.Typer(add_completion=False, pretty_exceptions_enable=False)
new_app = typer.Typer()
sheets_app = typer.Typer()
assets_app = typer.Typer()
workbooks_app = typer.Typer()
app.add_typer(new_app, name="new")
app.add_typer(sheets_app, name="sheets")
app.add_typer(assets_app, name="assets")
app.add_typer(workbooks_app, name="workbooks")


def emit(payload: Any, fmt: str = "text") -> None:
    if fmt == "json":
        typer.echo(json.dumps(payload, indent=2))
        return
    if isinstance(payload, dict):
        typer.echo(payload.get("message", json.dumps(payload, indent=2)))
    else:
        typer.echo(payload)


def fail(err: HerndonError, fmt: str) -> None:
    emit(err.to_dict() if fmt == "json" else err.message, fmt)
    raise typer.Exit(err.exit_code)


def save_sheet(path: Path, sheet: SheetSpec, dry_run: bool) -> None:
    if dry_run:
        return
    atomic_dump_json(path, sheet.model_dump(mode="json", exclude_none=True))


def save_workbook(path: Path, workbook: WorkbookSpec, dry_run: bool) -> None:
    if dry_run:
        return
    atomic_dump_json(path, workbook.model_dump(mode="json", exclude_none=True))


def _load_workbook_spec(workbook_json: str) -> tuple[Path, WorkbookSpec]:
    workbook_path = Path(workbook_json).resolve()
    return workbook_path, WorkbookSpec.model_validate(load_json(workbook_path))


@app.command()
def init(path: str = typer.Argument("."), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    target = Path(path).resolve()
    project_name = target.name
    structure = [
        target / ".herndon" / "themes",
        target / ".herndon" / "builds",
        target / ".herndon" / "cache",
        target / ".herndon" / "logs",
        target / "workbooks",
        target / "assets",
    ]
    if not dry_run:
        for entry in structure:
            entry.mkdir(parents=True, exist_ok=True)
        dump_json(target / ".herndon" / "config.json", ProjectConfig(version=1, project_name=project_name).model_dump(mode="json"))
    emit({"ok": True, "path": str(target), "message": f"Initialized project at {target}"}, format)


@new_app.command("workbook")
def new_workbook(name: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    project_root = require_project_root(Path.cwd())
    workbook_dir = project_root / "workbooks" / name
    workbook_path = workbook_dir / "workbook.json"
    workbook = WorkbookSpec(
        version=1,
        workbook_id=name,
        title=name.replace("_", " ").title(),
        theme=None,
        sheets=[],
        build={"output": f".herndon/builds/{name}/{name}.xlsx"},
    )
    if not dry_run:
        (workbook_dir / "sheets").mkdir(parents=True, exist_ok=True)
        save_workbook(workbook_path, workbook, dry_run=False)
    emit({"ok": True, "workbook_path": str(workbook_path), "message": f"Created workbook {workbook_path}"}, format)


@new_app.command("sheet")
def new_sheet(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    workbook_path = Path(workbook_json).resolve()
    loaded = load_workbook(workbook_path)
    sheet_path = workbook_path.parent / "sheets" / f"{len(loaded.spec.sheets)+1:03d}-{sheet_id}.json"
    sheet = SheetSpec(sheet_id=sheet_id, title=sheet_id.replace("_", " ").title())
    loaded.spec.sheets.append(str(sheet_path.relative_to(workbook_path.parent)))
    if not dry_run:
        save_sheet(sheet_path, sheet, dry_run=False)
        save_workbook(workbook_path, loaded.spec, dry_run=False)
    emit({"ok": True, "sheet_path": str(sheet_path), "message": f"Created sheet {sheet_path}"}, format)


@app.command()
def validate(target: str = typer.Argument(...), format: str = typer.Option("text", "--format"), project_root: Optional[str] = typer.Option(None, "--project-root")) -> None:
    root = Path(project_root).resolve() if project_root else None
    result = validate_path(Path(target), project_root=root)
    if format == "json":
        emit(result, format)
    else:
        if result["ok"]:
            typer.echo("Validation passed")
        else:
            for item in result["issues"]:
                typer.echo(f"{item['severity'].upper()} {item['code']}: {item['message']} ({item['path']})")
    if not result["ok"]:
        raise typer.Exit(1)


@app.command()
def inspect(target: str = typer.Argument(...), format: str = typer.Option("json", "--format")) -> None:
    target_path = Path(target).resolve()
    if target_path.name == "workbook.json":
        payload = inspect_workbook(load_workbook(target_path))
    else:
        loaded_sheet = load_sheet(target_path)
        payload = inspect_sheet(target_path, loaded_sheet.spec)
    emit(payload, format)


@app.command()
def render(workbook_path: str = typer.Argument(...), format: str = typer.Option("text", "--format")) -> None:
    manifest = render_workbook(load_workbook(Path(workbook_path).resolve()))
    emit({"ok": True, "manifest": manifest, "message": f"Rendered {manifest['output_path']}"}, format)


@app.command()
def themes(project_root: Optional[str] = typer.Option(None, "--project-root"), format: str = typer.Option("text", "--format")) -> None:
    root = require_project_root(Path.cwd(), explicit_root=Path(project_root) if project_root else None)
    theme_dir = root / ".herndon" / "themes"
    payload = {"ok": True, "themes": sorted(path.stem for path in theme_dir.glob("*.json"))}
    emit(payload, format)


@assets_app.command("list")
def assets_list(project_root: Optional[str] = typer.Option(None, "--project-root"), format: str = typer.Option("text", "--format")) -> None:
    root = require_project_root(Path.cwd(), explicit_root=Path(project_root) if project_root else None)
    payload = {"ok": True, "assets": sorted(str(path.relative_to(root)) for path in (root / "assets").rglob("*") if path.is_file())}
    emit(payload, format)


@assets_app.command("inspect")
def assets_inspect(asset_path: str = typer.Argument(...), format: str = typer.Option("text", "--format")) -> None:
    path = Path(asset_path).resolve()
    with Image.open(path) as image:
        payload = {"ok": True, "path": str(path), "width": image.width, "height": image.height, "format": image.format}
    emit(payload, format)


def _load_workbook_and_sheet(workbook_json: str, sheet_id: str):
    loaded = load_workbook(Path(workbook_json).resolve())
    for item in loaded.sheets:
        if item.spec.sheet_id == sheet_id:
            return loaded, item
    raise HerndonError("io_error", f"Sheet '{sheet_id}' not found")


@workbooks_app.command("set-title")
def workbook_set_title(
    workbook_json: str = typer.Argument(...),
    title: str = typer.Argument(...),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    workbook_path, workbook = _load_workbook_spec(workbook_json)
    workbook.title = title
    save_workbook(workbook_path, workbook, dry_run)
    emit({"ok": True, "message": f"Updated workbook title to '{title}'"}, format)


@workbooks_app.command("set-theme")
def workbook_set_theme(
    workbook_json: str = typer.Argument(...),
    theme: str = typer.Argument(...),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    workbook_path, workbook = _load_workbook_spec(workbook_json)
    workbook.theme = theme
    save_workbook(workbook_path, workbook, dry_run)
    emit({"ok": True, "message": f"Updated workbook theme to '{theme}'"}, format)


@workbooks_app.command("set-output")
def workbook_set_output(
    workbook_json: str = typer.Argument(...),
    output: str = typer.Argument(...),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    workbook_path, workbook = _load_workbook_spec(workbook_json)
    workbook.build.output = output
    save_workbook(workbook_path, workbook, dry_run)
    emit({"ok": True, "message": f"Updated workbook output to '{output}'"}, format)


@sheets_app.command("list")
def sheets_list(workbook_json: str = typer.Argument(...), format: str = typer.Option("text", "--format")) -> None:
    loaded = load_workbook(Path(workbook_json).resolve())
    payload = {
        "ok": True,
        "sheets": [
            {"sheet_id": sheet.spec.sheet_id, "title": sheet.spec.title, "path": str(sheet.path)}
            for sheet in loaded.sheets
        ],
    }
    emit(payload, format)


@sheets_app.command("add")
def sheets_add(
    workbook_json: str = typer.Argument(...),
    sheet: str = typer.Option(..., "--sheet"),
    after: Optional[str] = typer.Option(None, "--after"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    workbook_path = Path(workbook_json).resolve()
    workbook = WorkbookSpec.model_validate(load_json(workbook_path))
    sheet_path = Path(sheet).resolve()
    rel = str(sheet_path.relative_to(workbook_path.parent))
    if rel in workbook.sheets:
        raise HerndonError("usage_error", "Sheet is already attached to workbook")
    if after is None:
        workbook.sheets.append(rel)
    else:
        loaded = load_workbook(workbook_path)
        ordered_ids = [item.spec.sheet_id for item in loaded.sheets]
        if after not in ordered_ids:
            raise HerndonError("usage_error", f"Unknown sheet_id '{after}'")
        insert_at = ordered_ids.index(after) + 1
        workbook.sheets.insert(insert_at, rel)
    save_workbook(workbook_path, workbook, dry_run)
    emit({"ok": True, "message": "Sheet added to workbook"}, format)


@sheets_app.command("remove")
def sheets_remove(
    workbook_json: str = typer.Argument(...),
    sheet_id: str = typer.Argument(...),
    delete_files: bool = typer.Option(False, "--delete-files"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    loaded, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    rel = str(target.path.relative_to(loaded.path.parent))
    loaded.spec.sheets = [entry for entry in loaded.spec.sheets if entry != rel]
    save_workbook(loaded.path, loaded.spec, dry_run)
    if delete_files and not dry_run:
        target.path.unlink(missing_ok=True)
    emit({"ok": True, "message": f"Removed sheet '{sheet_id}'"}, format)


@sheets_app.command("rename")
def sheets_rename(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), new_sheet_id: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    loaded, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    old_path = target.path
    new_path = old_path.with_name(old_path.name.replace(f"-{sheet_id}.json", f"-{new_sheet_id}.json"))
    target.spec.sheet_id = new_sheet_id
    target.spec.title = new_sheet_id.replace("_", " ").title()
    loaded.spec.sheets = [str(new_path.relative_to(loaded.path.parent)) if entry == str(old_path.relative_to(loaded.path.parent)) else entry for entry in loaded.spec.sheets]
    if not dry_run:
        save_sheet(new_path, target.spec, dry_run=False)
        if new_path != old_path:
            old_path.unlink(missing_ok=True)
        save_workbook(loaded.path, loaded.spec, dry_run=False)
    emit({"ok": True, "message": f"Renamed sheet '{sheet_id}' to '{new_sheet_id}'"}, format)


@sheets_app.command("duplicate")
def sheets_duplicate(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), new_sheet_id: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    loaded, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    new_path = target.path.with_name(target.path.name.replace(f"-{sheet_id}.json", f"-{new_sheet_id}.json"))
    new_sheet = target.spec.model_copy(deep=True)
    new_sheet.sheet_id = new_sheet_id
    new_sheet.title = new_sheet_id.replace("_", " ").title()
    loaded.spec.sheets.append(str(new_path.relative_to(loaded.path.parent)))
    if not dry_run:
        save_sheet(new_path, new_sheet, dry_run=False)
        save_workbook(loaded.path, loaded.spec, dry_run=False)
    emit({"ok": True, "message": f"Duplicated sheet '{sheet_id}' as '{new_sheet_id}'"}, format)


@sheets_app.command("move")
def sheets_move(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), after: str = typer.Option(..., "--after"), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    loaded, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    rel = str(target.path.relative_to(loaded.path.parent))
    ordered = [entry for entry in loaded.spec.sheets if entry != rel]
    id_map = {sheet.spec.sheet_id: str(sheet.path.relative_to(loaded.path.parent)) for sheet in loaded.sheets}
    if after not in id_map:
        raise HerndonError("usage_error", f"Unknown sheet_id '{after}'")
    insert_at = ordered.index(id_map[after]) + 1
    ordered.insert(insert_at, rel)
    loaded.spec.sheets = ordered
    save_workbook(loaded.path, loaded.spec, dry_run)
    emit({"ok": True, "message": f"Moved sheet '{sheet_id}'"}, format)


@sheets_app.command("set-cell")
def set_cell(
    workbook_json: str = typer.Argument(...),
    sheet_id: str = typer.Argument(...),
    cell: str = typer.Argument(...),
    value: Optional[str] = typer.Option(None, "--value"),
    formula: Optional[str] = typer.Option(None, "--formula"),
    style: Optional[str] = typer.Option(None, "--style"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    loaded, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.cells = [entry for entry in target.spec.cells if entry.cell != cell]
    payload = {"cell": cell, "style": style}
    if formula is not None:
        payload["formula"] = formula
    else:
        payload["value"] = parse_cli_value(value) if value is not None else None
    target.spec.cells.append(CellSpec.model_validate(payload))
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Updated cell {cell}"}, format)


@sheets_app.command("set-range")
def set_range(
    workbook_json: str = typer.Argument(...),
    sheet_id: str = typer.Argument(...),
    anchor: str = typer.Argument(...),
    data_json: str = typer.Option(..., "--data-json"),
    row_styles: str = typer.Option("{}", "--row-styles"),
    col_styles: str = typer.Option("{}", "--col-styles"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.ranges = [item for item in target.spec.ranges if item.anchor != anchor]
    target.spec.ranges.append(RangeSpec.model_validate({"anchor": anchor, "data": json.loads(data_json), "row_styles": json.loads(row_styles), "col_styles": json.loads(col_styles)}))
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Updated range at {anchor}"}, format)


@sheets_app.command("add-table")
def add_table(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), table_id: str = typer.Argument(...), ref: str = typer.Option(..., "--ref"), name: str = typer.Option(..., "--name"), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.tables = [item for item in target.spec.tables if item.table_id != table_id]
    target.spec.tables.append(TableSpec.model_validate({"table_id": table_id, "ref": ref, "name": name}))
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Added table '{table_id}'"}, format)


@sheets_app.command("add-chart")
def add_chart(
    workbook_json: str = typer.Argument(...),
    sheet_id: str = typer.Argument(...),
    chart_id: str = typer.Argument(...),
    type: str = typer.Option(..., "--type"),
    anchor: str = typer.Option(..., "--anchor"),
    series_json: str = typer.Option(..., "--series-json"),
    title: Optional[str] = typer.Option(None, "--title"),
    stacked: bool = typer.Option(False, "--stacked"),
    percent_stacked: bool = typer.Option(False, "--percent-stacked"),
    legend_position: Optional[str] = typer.Option(None, "--legend-position"),
    show_data_labels: bool = typer.Option(False, "--show-data-labels"),
    show_percent_labels: bool = typer.Option(False, "--show-percent-labels"),
    x_axis_title: Optional[str] = typer.Option(None, "--x-axis-title"),
    y_axis_title: Optional[str] = typer.Option(None, "--y-axis-title"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.charts = [item for item in target.spec.charts if item.chart_id != chart_id]
    target.spec.charts.append(
        ChartSpec.model_validate(
            {
                "chart_id": chart_id,
                "chart_type": type,
                "anchor": anchor,
                "series": json.loads(series_json),
                "title": title,
                "stacked": stacked,
                "percent_stacked": percent_stacked,
                "legend_position": legend_position,
                "show_data_labels": show_data_labels,
                "show_percent_labels": show_percent_labels,
                "x_axis_title": x_axis_title,
                "y_axis_title": y_axis_title,
            }
        )
    )
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Added chart '{chart_id}'"}, format)


@sheets_app.command("update-chart")
def update_chart(
    workbook_json: str = typer.Argument(...),
    sheet_id: str = typer.Argument(...),
    chart_id: str = typer.Argument(...),
    title: Optional[str] = typer.Option(None, "--title"),
    show_legend: bool = typer.Option(False, "--show-legend"),
    legend_position: Optional[str] = typer.Option(None, "--legend-position"),
    show_data_labels: bool = typer.Option(False, "--show-data-labels"),
    show_percent_labels: bool = typer.Option(False, "--show-percent-labels"),
    x_axis_title: Optional[str] = typer.Option(None, "--x-axis-title"),
    y_axis_title: Optional[str] = typer.Option(None, "--y-axis-title"),
    format: str = typer.Option("text", "--format"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    for chart in target.spec.charts:
        if chart.chart_id == chart_id:
            if title is not None:
                chart.title = title
            if show_legend:
                chart.show_legend = True
            if legend_position is not None:
                chart.legend_position = legend_position
            if show_data_labels:
                chart.show_data_labels = True
            if show_percent_labels:
                chart.show_percent_labels = True
            if x_axis_title is not None:
                chart.x_axis_title = x_axis_title
            if y_axis_title is not None:
                chart.y_axis_title = y_axis_title
            save_sheet(target.path, target.spec, dry_run)
            emit({"ok": True, "message": f"Updated chart '{chart_id}'"}, format)
            return
    raise HerndonError("io_error", f"Chart '{chart_id}' not found")


@sheets_app.command("remove-element")
def remove_element(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), element_id: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.tables = [item for item in target.spec.tables if item.table_id != element_id]
    target.spec.charts = [item for item in target.spec.charts if item.chart_id != element_id]
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Removed element '{element_id}'"}, format)


@sheets_app.command("set-merge")
def set_merge(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), range_ref: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    if range_ref not in target.spec.merges:
        target.spec.merges.append(range_ref)
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Added merge '{range_ref}'"}, format)


@sheets_app.command("clear-merge")
def clear_merge(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), range_ref: str = typer.Argument(...), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.merges = [item for item in target.spec.merges if item != range_ref]
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Removed merge '{range_ref}'"}, format)


@sheets_app.command("freeze")
def freeze(workbook_json: str = typer.Argument(...), sheet_id: str = typer.Argument(...), rows: int = typer.Option(0, "--rows"), cols: int = typer.Option(0, "--cols"), format: str = typer.Option("text", "--format"), dry_run: bool = typer.Option(False, "--dry-run")) -> None:
    _, target = _load_workbook_and_sheet(workbook_json, sheet_id)
    target.spec.freeze_rows = rows
    target.spec.freeze_cols = cols
    save_sheet(target.path, target.spec, dry_run)
    emit({"ok": True, "message": f"Updated freeze panes for '{sheet_id}'"}, format)


def main() -> None:
    try:
        app()
    except ValidationError as exc:
        fail(HerndonError("schema_error", "Schema validation failed", details=[{"message": str(exc)}]), "json" if "--format" in __import__("sys").argv and "json" in __import__("sys").argv else "text")
    except HerndonError as exc:
        fail(exc, "json" if "--format" in __import__("sys").argv and "json" in __import__("sys").argv else "text")


if __name__ == "__main__":
    main()
