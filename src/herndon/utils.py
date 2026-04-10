from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

CELL_RE = re.compile(r"^[A-Z]+[1-9][0-9]*$")
RANGE_RE = re.compile(r"^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$")
HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
SHEET_REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_]+))!")


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def dump_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, indent=2, sort_keys=False) + "\n"
    path.write_text(text, encoding="utf-8")


def atomic_dump_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, delete=False) as handle:
        handle.write(json.dumps(payload, indent=2, sort_keys=False) + "\n")
        temp_path = Path(handle.name)
    temp_path.replace(path)


def is_cell_address(value: str) -> bool:
    return bool(CELL_RE.fullmatch(value))


def is_range_address(value: str) -> bool:
    return bool(RANGE_RE.fullmatch(value))


def is_hex_color(value: str) -> bool:
    return bool(HEX_RE.fullmatch(value))


def normalize_excel_color(value: str) -> str:
    return value.replace("#", "").upper()


def parse_cli_value(raw: str) -> Any:
    lowered = raw.lower()
    if lowered == "null":
        return None
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        if "." in raw:
            return float(raw)
        return int(raw)
    except ValueError:
        return raw


def looks_like_iso_date(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return len(value) == 10


def timestamp_utc() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def sheet_refs_from_formula(formula: str) -> set[str]:
    refs: set[str] = set()
    for match in SHEET_REF_RE.finditer(formula):
        refs.add(match.group(1) or match.group(2))
    return refs


def find_project_root(start: Path) -> Path | None:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / ".herndon" / "config.json").exists():
            return candidate
    return None


def make_relative(path: Path, root: Path) -> str:
    return str(path.resolve().relative_to(root.resolve()))
