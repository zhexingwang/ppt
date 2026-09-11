#!/usr/bin/env python3
"""発明開示書作成手順（訂正反映版）を生成する。"""

from __future__ import annotations

import argparse
from pathlib import Path

from infographic.invention_disclosure import build_deck

DEFAULT_OUT = Path("decks/05_Copilot発明開示書/Copilot発明開示書_訂正反映.pptx")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    prs = build_deck()
    prs.save(args.output)
    print(f"Generated: {args.output.resolve()}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
