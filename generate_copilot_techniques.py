#!/usr/bin/env python3
"""Copilotとのやり取りのコツ（発明開示書マニュアルの補足）を生成する。"""

from __future__ import annotations

import argparse
from pathlib import Path

from infographic.copilot_techniques import build_deck

DEFAULT_OUT = Path("decks/06_Copilot使いこなし/Copilot使いこなしのコツ.pptx")


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
