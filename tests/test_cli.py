from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from typer.testing import CliRunner

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from herndon.cli import app


runner = CliRunner()


def test_init_new_workbook_and_sheet_flow(tmp_path: Path) -> None:
    result = runner.invoke(app, ["init", str(tmp_path)])
    assert result.exit_code == 0

    with runner.isolated_filesystem(temp_dir=tmp_path):
        pass

    cwd = Path.cwd()
    try:
        import os

        os.chdir(tmp_path)
        result = runner.invoke(app, ["new", "workbook", "q2_report"])
        assert result.exit_code == 0
        result = runner.invoke(app, ["new", "sheet", str(tmp_path / "workbooks" / "q2_report" / "workbook.json"), "summary"])
        assert result.exit_code == 0
        workbook = json.loads((tmp_path / "workbooks" / "q2_report" / "workbook.json").read_text())
        assert workbook["sheets"] == ["sheets/001-summary.json"]
    finally:
        os.chdir(cwd)


def test_set_cell_validate_and_inspect_json(tmp_path: Path) -> None:
    runner.invoke(app, ["init", str(tmp_path)])
    cwd = Path.cwd()
    try:
        os.chdir(tmp_path)
        workbook_path = tmp_path / "workbooks" / "demo" / "workbook.json"
        runner.invoke(app, ["new", "workbook", "demo"])
        runner.invoke(app, ["new", "sheet", str(workbook_path), "summary"])
        runner.invoke(app, ["sheets", "set-cell", str(workbook_path), "summary", "A1", "--value", "Revenue", "--style", "header"])
        issues = runner.invoke(app, ["validate", str(workbook_path), "--format", "json"])
        assert issues.exit_code == 1
        payload = json.loads(issues.stdout)
        assert payload["ok"] is False
        inspect_result = runner.invoke(app, ["inspect", str(workbook_path), "--format", "json"])
        assert inspect_result.exit_code == 0
        inspect_payload = json.loads(inspect_result.stdout)
        assert inspect_payload["sheets"][0]["cells"][0]["cell"] == "A1"
    finally:
        os.chdir(cwd)


def test_workbook_metadata_commands_and_inspect_golden(tmp_path: Path) -> None:
    runner.invoke(app, ["init", str(tmp_path)])
    cwd = Path.cwd()
    try:
        os.chdir(tmp_path)
        workbook_path = tmp_path / "workbooks" / "demo" / "workbook.json"
        runner.invoke(app, ["new", "workbook", "demo"])
        runner.invoke(app, ["new", "sheet", str(workbook_path), "summary"])
        runner.invoke(app, ["workbooks", "set-title", str(workbook_path), "Demo Finance Book"])
        runner.invoke(app, ["workbooks", "set-theme", str(workbook_path), "brand"])
        runner.invoke(app, ["workbooks", "set-output", str(workbook_path), ".herndon/builds/demo/finance-demo.xlsx"])
        theme_path = tmp_path / ".herndon" / "themes" / "brand.json"
        theme_path.write_text(json.dumps({"name": "brand", "styles": {"header": {"bold": True}}}, indent=2))
        runner.invoke(app, ["sheets", "set-cell", str(workbook_path), "summary", "A1", "--value", "Revenue", "--style", "header"])

        inspect_result = runner.invoke(app, ["inspect", str(workbook_path), "--format", "json"])
        assert inspect_result.exit_code == 0
        payload = json.loads(inspect_result.stdout)
        payload["project_root"] = "__PROJECT_ROOT__"
        payload["sheets"][0]["path"] = "__PROJECT_ROOT__/workbooks/demo/sheets/001-summary.json"

        expected = json.loads((Path(__file__).parent / "fixtures" / "inspect_workbook_expected.json").read_text())
        assert payload == expected
    finally:
        os.chdir(cwd)
