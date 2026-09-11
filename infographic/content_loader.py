"""JSONからデッキ内容を読み込む（カスタム用）。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pptx import Presentation

from .slides import (
    add_comparison_slide,
    add_kpi_slide,
    add_matrix_slide,
    add_pillars_slide,
    add_process_slide,
    add_takeaways_slide,
    add_timeline_slide,
    add_title_slide,
)
from .theme import Theme


BUILDERS = {
    "title": add_title_slide,
    "kpi": add_kpi_slide,
    "process": add_process_slide,
    "comparison": add_comparison_slide,
    "timeline": add_timeline_slide,
    "pillars": add_pillars_slide,
    "matrix": add_matrix_slide,
    "takeaways": add_takeaways_slide,
}


def build_from_json(path: str | Path) -> Presentation:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT

    for slide in data.get("slides", []):
        kind = slide.get("type")
        if kind not in BUILDERS:
            raise ValueError(f"Unknown slide type: {kind}")
        payload = {k: v for k, v in slide.items() if k != "type"}
        BUILDERS[kind](prs, **payload)
    return prs
