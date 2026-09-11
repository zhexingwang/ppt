#!/usr/bin/env python3
"""競争吸着 v4：反応一覧・定数・VE出口貢献デッキを生成する。"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from infographic.ve_contribution import build_deck

DEFAULT_OUT = Path("decks/04_反応一覧_VE貢献_v4/反応一覧_定数_VE出口貢献_v4.pptx")
IMAGE_CANDIDATES = {
    "ve_reaction_map.png": Path("/opt/cursor/artifacts/assets/ve_reaction_map.png"),
    "ve_causal_chain.png": Path("/opt/cursor/artifacts/assets/ve_causal_chain.png"),
}


def sync_assets():
    dest = Path("assets")
    dest.mkdir(parents=True, exist_ok=True)
    for name, src in IMAGE_CANDIDATES.items():
        if src.exists():
            shutil.copy2(src, dest / name)


def main():
    parser = argparse.ArgumentParser(description="VE貢献インフォグラフィックPPTを生成")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()

    sync_assets()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    prs = build_deck()
    prs.save(args.output)
    print(f"Generated: {args.output.resolve()}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
