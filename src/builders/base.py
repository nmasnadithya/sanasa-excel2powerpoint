"""Builder protocol — v1 ships TemplateBuilder; future builders implement same."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Protocol

from src.excel_reader import ExcelReader
from src.slide_specs import SlideSpec


class Builder(Protocol):
    def build(
        self,
        specs: list[SlideSpec],
        reader: ExcelReader,
        target: date,
        output_path: Path,
    ) -> None: ...
