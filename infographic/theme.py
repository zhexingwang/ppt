"""共通テーマ（色・余白・フォント）。"""

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


class Theme:
    """ワイド画面(16:9)向けのビジネス配色。"""

    # スライドサイズ
    SLIDE_WIDTH = Inches(13.333)
    SLIDE_HEIGHT = Inches(7.5)

    # カラーパレット
    PRIMARY = RGBColor(0x1F, 0x4E, 0x79)  # ディープブルー
    PRIMARY_LIGHT = RGBColor(0x2E, 0x75, 0xB6)
    ACCENT = RGBColor(0xED, 0x7D, 0x31)  # オレンジ
    ACCENT_2 = RGBColor(0x70, 0xAD, 0x47)  # グリーン
    ACCENT_3 = RGBColor(0x5B, 0x9B, 0xD5)  # ライトブルー
    ACCENT_4 = RGBColor(0xA5, 0xA5, 0xA5)  # グレー
    BG = RGBColor(0xF7, 0xF9, 0xFC)
    CARD = RGBColor(0xFF, 0xFF, 0xFF)
    TEXT = RGBColor(0x2D, 0x2D, 0x2D)
    TEXT_MUTED = RGBColor(0x66, 0x66, 0x66)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    DIVIDER = RGBColor(0xD6, 0xDE, 0xE8)

    # カード配色セット
    CARD_COLORS = [
        RGBColor(0x1F, 0x4E, 0x79),
        RGBColor(0x2E, 0x75, 0xB6),
        RGBColor(0x5B, 0x9B, 0xD5),
        RGBColor(0xED, 0x7D, 0x31),
        RGBColor(0x70, 0xAD, 0x47),
    ]

    # フォント（Windowsでは "Yu Gothic" / "Meiryo" に変更推奨）
    FONT_JP = "Noto Sans CJK JP"
    FONT_EN = "Calibri"

    # 余白
    MARGIN_X = Inches(0.5)
    MARGIN_Y = Inches(0.4)

    # タイポグラフィ
    TITLE_SIZE = Pt(32)
    SUBTITLE_SIZE = Pt(16)
    BODY_SIZE = Pt(14)
    SMALL_SIZE = Pt(11)
    NUMBER_SIZE = Pt(36)
