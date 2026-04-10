from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from .utils import is_cell_address, is_hex_color, is_range_address


class HerndonBaseModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class CellSpec(HerndonBaseModel):
    cell: str
    value: str | int | float | bool | None = None
    formula: str | None = None
    style: str | None = None
    image_path: str | None = None
    w: float | None = None
    h: float | None = None

    @field_validator("cell")
    @classmethod
    def validate_cell(cls, value: str) -> str:
        if not is_cell_address(value):
            raise ValueError("invalid cell address")
        return value

    @model_validator(mode="after")
    def validate_mode(self) -> "CellSpec":
        if self.formula is not None and self.value is not None:
            raise ValueError("value and formula are mutually exclusive")
        if self.formula is not None and not self.formula.startswith("="):
            raise ValueError("formula must start with '='")
        if self.image_path is not None and (self.w is None or self.h is None):
            raise ValueError("image placements require w and h")
        return self


class RangeSpec(HerndonBaseModel):
    anchor: str
    data: list[list[Any]]
    row_styles: dict[str, str] = Field(default_factory=dict)
    col_styles: dict[str, str] = Field(default_factory=dict)

    @field_validator("anchor")
    @classmethod
    def validate_anchor(cls, value: str) -> str:
        if not is_cell_address(value):
            raise ValueError("invalid anchor cell")
        return value


class TableSpec(HerndonBaseModel):
    table_id: str
    name: str
    ref: str
    header_row: bool = True
    auto_filter: bool = True
    style: str | None = "TableStyleMedium2"

    @field_validator("ref")
    @classmethod
    def validate_ref(cls, value: str) -> str:
        if not is_range_address(value):
            raise ValueError("invalid table range")
        return value


class ChartSeriesSpec(HerndonBaseModel):
    label: str
    values: str
    categories: str | None = None
    color: str | None = None

    @field_validator("color")
    @classmethod
    def validate_color(cls, value: str | None) -> str | None:
        if value is not None and not is_hex_color(value):
            raise ValueError("invalid hex color")
        return value


class ChartSpec(HerndonBaseModel):
    chart_id: str
    chart_type: Literal["bar", "column", "line", "pie", "scatter"]
    title: str | None = None
    anchor: str
    w: float = 8.0
    h: float = 5.0
    series: list[ChartSeriesSpec]
    stacked: bool = False
    percent_stacked: bool = False
    show_legend: bool = True
    legend_position: Literal["l", "r", "t", "b", "tr"] | None = None
    show_data_labels: bool = False
    show_percent_labels: bool = False
    x_axis_title: str | None = None
    y_axis_title: str | None = None
    value_format: str | None = None

    @field_validator("anchor")
    @classmethod
    def validate_anchor(cls, value: str) -> str:
        if not is_cell_address(value):
            raise ValueError("invalid anchor cell")
        return value

    @model_validator(mode="after")
    def validate_chart_mode(self) -> "ChartSpec":
        if self.stacked and self.percent_stacked:
            raise ValueError("stacked and percent_stacked are mutually exclusive")
        return self


class SheetSpec(HerndonBaseModel):
    sheet_id: str
    title: str
    tab_color: str | None = None
    freeze_rows: int = 0
    freeze_cols: int = 0
    zoom: int = 100
    column_widths: dict[str, float] = Field(default_factory=dict)
    row_heights: dict[str, float] = Field(default_factory=dict)
    cells: list[CellSpec] = Field(default_factory=list)
    ranges: list[RangeSpec] = Field(default_factory=list)
    merges: list[str] = Field(default_factory=list)
    tables: list[TableSpec] = Field(default_factory=list)
    charts: list[ChartSpec] = Field(default_factory=list)

    @field_validator("tab_color")
    @classmethod
    def validate_tab_color(cls, value: str | None) -> str | None:
        if value is not None and not is_hex_color(value):
            raise ValueError("invalid hex color")
        return value

    @field_validator("merges")
    @classmethod
    def validate_merges(cls, value: list[str]) -> list[str]:
        for item in value:
            if not is_range_address(item):
                raise ValueError(f"invalid merge range: {item}")
        return value


class BuildSpec(HerndonBaseModel):
    output: str


class WorkbookSpec(HerndonBaseModel):
    version: int = 1
    workbook_id: str
    title: str
    theme: str | None = None
    sheets: list[str] = Field(default_factory=list)
    build: BuildSpec


class ThemeStyle(HerndonBaseModel):
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    fill: str | None = None
    number_format: str | None = None
    alignment: Literal["left", "center", "right", "general"] | None = None
    vertical_alignment: Literal["top", "middle", "bottom"] | None = None
    wrap_text: bool | None = None
    border_top: str | None = None
    border_bottom: str | None = None
    border_left: str | None = None
    border_right: str | None = None
    border_color: str | None = None


class ThemeSpec(HerndonBaseModel):
    name: str
    colors: dict[str, str] = Field(default_factory=dict)
    fonts: dict[str, str] = Field(default_factory=dict)
    styles: dict[str, ThemeStyle] = Field(default_factory=dict)

    @field_validator("colors")
    @classmethod
    def validate_colors(cls, value: dict[str, str]) -> dict[str, str]:
        for color in value.values():
            if not is_hex_color(color):
                raise ValueError("invalid theme color")
        return value


class ProjectConfig(HerndonBaseModel):
    version: int = 1
    project_name: str


class LoadedSheet(BaseModel):
    path: Path
    spec: SheetSpec

    model_config = ConfigDict(arbitrary_types_allowed=True)


class LoadedWorkbook(BaseModel):
    project_root: Path
    path: Path
    spec: WorkbookSpec
    sheets: list[LoadedSheet]
    theme: ThemeSpec | None = None

    model_config = ConfigDict(arbitrary_types_allowed=True)
