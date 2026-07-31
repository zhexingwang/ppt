#!/usr/bin/env python3
"""自己紹介インフォグラフィック（1枚）を生成する。"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from pptx import Presentation

from infographic.self_intro import add_self_intro_image_slide, add_self_intro_slide
from infographic.theme import Theme


DEFAULT_IMAGE_CANDIDATES = [
    Path("assets/self_intro_infographic.png"),
    Path("/opt/cursor/artifacts/assets/self_intro_infographic.png"),
]


def resolve_image(explicit: Path | None) -> Path | None:
    if explicit:
        return explicit if explicit.exists() else None
    for candidate in DEFAULT_IMAGE_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def main():
    parser = argparse.ArgumentParser(description="自己紹介スライドを生成します")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("output/self_intro.pptx"),
        help="出力 pptx パス",
    )
    parser.add_argument(
        "--image",
        type=Path,
        default=None,
        help="全面配置するインフォグラフィック画像（省略時は assets を自動探索）",
    )
    parser.add_argument(
        "--editable-only",
        action="store_true",
        help="編集可能版のみ生成（画像スライドを含めない）",
    )
    args = parser.parse_args()

    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT

    image_path = None if args.editable_only else resolve_image(args.image)
    if image_path:
        # リポジトリ内へコピーして再現性を確保
        assets_dir = Path("assets")
        assets_dir.mkdir(parents=True, exist_ok=True)
        local_image = assets_dir / "self_intro_infographic.png"
        if image_path.resolve() != local_image.resolve():
            shutil.copy2(image_path, local_image)
        add_self_intro_image_slide(prs, local_image)

    add_self_intro_slide(prs)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.output)
    print(f"Generated: {args.output.resolve()}")
    print(f"Slides: {len(prs.slides)}")
    if image_path:
        print(f"Image slide: {image_path}")
    print("Editable slide: included")


if __name__ == "__main__":
    main()
