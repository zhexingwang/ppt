"""Copilotで発明開示書を作る手順。"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .shapes import add_circle, add_rect, add_right_arrow, add_text_box, header_bar, set_paragraph_text
from .slides import add_title_slide, blank_slide
from .theme import Theme


NAVY = Theme.PRIMARY
ASK = RGBColor(0x1F, 0x4E, 0x79)
ANS = RGBColor(0x2E, 0x75, 0xB6)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
RED = RGBColor(0xC0, 0x39, 0x2B)
CREAM = RGBColor(0xFD, 0xF6, 0xE3)
GOAL_BG = RGBColor(0xE8, 0xF1, 0xFA)
PS_REF = "https://www.panasonic.com/jp/business/its/patentsquare.html"


def _footer(slide, page: str):
    add_text_box(
        slide, Inches(0.45), Inches(7.18), Inches(8.5), Inches(0.26),
        "Confidential(Yokogawa)", size=Pt(10), color=Theme.TEXT_MUTED,
    )
    add_text_box(
        slide, Inches(11.6), Inches(7.18), Inches(1.2), Inches(0.26),
        page, size=Pt(10), color=Theme.TEXT_MUTED, align=PP_ALIGN.RIGHT,
    )


def _paras(slide, left, top, width, height, text: str, *, size=Pt(11), color=Theme.TEXT, bold=False, space_after=Pt(3)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf.vertical_anchor = MSO_ANCHOR.TOP
    except Exception:
        pass
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_paragraph_text(p, line, size=size, color=color, bold=bold)
        p.space_after = space_after
        try:
            p.line_spacing = 1.05
        except Exception:
            pass
    return box


def _qa(slide, left, top, width, height, question: str, answer: str):
    add_rect(slide, left, top, width, height, Theme.CARD, corner=True)
    add_rect(slide, left, top, Inches(0.12), height, ASK)
    add_text_box(
        slide, left + Inches(0.25), top + Inches(0.08), width - Inches(0.4), Inches(0.24),
        "あなた", size=Pt(11), color=ASK, bold=True,
    )
    _paras(
        slide, left + Inches(0.25), top + Inches(0.30), width - Inches(0.4), Inches(0.62),
        question, size=Pt(11), color=Theme.TEXT, space_after=Pt(0),
    )
    add_rect(slide, left + Inches(0.2), top + Inches(0.96), width - Inches(0.4), Inches(0.015), Theme.DIVIDER)
    add_text_box(
        slide, left + Inches(0.25), top + Inches(1.02), width - Inches(0.4), Inches(0.24),
        "回答例（① 60点の素案）　※3枚目の①。人と直す前。Copilotでも同じ聞き方でよい。",
        size=Pt(11), color=ANS, bold=True,
    )
    _paras(
        slide, left + Inches(0.25), top + Inches(1.28), width - Inches(0.45), height - Inches(1.40),
        answer, size=Pt(10), color=Theme.TEXT, space_after=Pt(2),
    )


def _progress(slide, current: int):
    labels = ["0", "1", "2", "3", "4", "5"]
    x0 = Inches(9.15)
    y = Inches(0.32)
    add_text_box(slide, Inches(8.15), y, Inches(0.95), Inches(0.38), "いまここ", size=Pt(11), color=Theme.TEXT_MUTED, anchor=MSO_ANCHOR.MIDDLE)
    for i, lab in enumerate(labels):
        x = x0 + Inches(0.62) * i
        on = i == current
        add_circle(slide, x, y, Inches(0.38), NAVY if on else Theme.DIVIDER)
        add_text_box(
            slide, x, y + Inches(0.02), Inches(0.38), Inches(0.34), lab,
            size=Pt(11), color=Theme.WHITE if on else Theme.TEXT_MUTED, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )


def _flow_row(slide, items, top, *, left=None, box_w=None, box_h=None, gap=None):
    left = Inches(0.5) if left is None else left
    box_w = Inches(2.15) if box_w is None else box_w
    box_h = Inches(0.72) if box_h is None else box_h
    gap = Inches(0.42) if gap is None else gap
    for i, t in enumerate(items):
        x = left + (box_w + gap) * i
        add_rect(slide, x, top, box_w, box_h, NAVY, corner=True)
        add_text_box(
            slide, x + Inches(0.06), top + Inches(0.08), box_w - Inches(0.12), box_h - Inches(0.14),
            t, size=Pt(12), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        if i < len(items) - 1:
            add_right_arrow(
                slide, x + box_w + Inches(0.04), top + Inches(0.22), Inches(0.34), Inches(0.28), GOLD,
            )


def add_purpose_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "この資料で伝えたいこと", "発明開示書の素案を、Copilotで速く・抜けなく作る")
    cards = [
        ("誰向け", "初めて開示書を書く人\nAIで知財業務を進めたい人"),
        ("何が変わる", "数日〜1週間 → 数時間で素案\n中身の責任は発明者自身"),
        ("進め方", "作業順はStep 0〜5\n先行特許は2回見る"),
    ]
    for i, (t, b) in enumerate(cards):
        x = Inches(0.5) + Inches(4.2) * i
        add_rect(slide, x, Inches(1.5), Inches(3.95), Inches(4.6), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.5), Inches(3.95), Inches(0.7), NAVY)
        add_text_box(slide, x, Inches(1.6), Inches(3.95), Inches(0.5), t, size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, x + Inches(0.25), Inches(2.5), Inches(3.45), Inches(3.2), b, size=Pt(16), color=Theme.TEXT)
    _footer(slide, page)
    return slide


def add_guardrail_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "使う前の約束", "入力してよい情報と、AIの位置づけ")
    rows = [
        ("開示書作成中（発明のタネ）", "Restricted", "Copilotへ入力可"),
        ("ANAQUA登録後", "Confidential", "関係者のみ。Copilotへ入れない"),
        ("共同研究の情報", "Confidential", "生成AIへ原則禁止"),
    ]
    y = Inches(1.35)
    for title, cls, note in rows:
        add_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.85), Theme.CARD, corner=True)
        add_text_box(slide, Inches(0.7), y + Inches(0.2), Inches(5.5), Inches(0.5), title, size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, Inches(6.4), y + Inches(0.2), Inches(2.4), Inches(0.5), cls, size=Pt(14), color=RED, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, Inches(8.8), y + Inches(0.2), Inches(3.7), Inches(0.5), note, size=Pt(14), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.98)
    add_rect(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.3), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.75), Inches(4.65), Inches(11.8), Inches(0.45), "AIは支援者。開示書の名義はあなたです。", size=Pt(18), color=NAVY, bold=True)
    add_text_box(
        slide, Inches(0.75), Inches(5.2), Inches(11.8), Inches(1.35),
        "① Copilotで60点の素案  →  ② 人と議論して直す  →  ③ もう一度Copilotで整える\n"
        "各Stepの「回答例」は①。「仕上げ例」は③。出力をそのまま貼らない。",
        size=Pt(15), color=Theme.TEXT,
    )
    _footer(slide, page)
    return slide


def add_tools_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "使う道具", "文章はCopilot。PatentSQUAREは、クレーム案のあとの衝突確認で使う")

    add_rect(slide, Inches(0.5), Inches(1.45), Inches(8.05), Inches(5.35), Theme.CARD, corner=True)
    add_rect(slide, Inches(0.5), Inches(1.45), Inches(8.05), Inches(0.85), NAVY)
    add_text_box(
        slide, Inches(0.5), Inches(1.58), Inches(8.05), Inches(0.58),
        "まずはこれ　M365 Copilot", size=Pt(22), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER,
    )
    add_text_box(
        slide, Inches(0.85), Inches(2.55), Inches(7.35), Inches(3.85),
        "・文章の下書き・整理・用語そろえ\n"
        "・Step 1の先行特許調査（背景と、クレームの方向性）\n"
        "・図（Graphviz）も、社外ツールを使わずここで作る\n\n"
        "自然な言葉で依頼する。出力はそのまま貼らない。",
        size=Pt(16), color=Theme.TEXT,
    )

    add_rect(slide, Inches(8.75), Inches(1.45), Inches(4.05), Inches(5.35), Theme.CARD, corner=True)
    add_rect(slide, Inches(8.75), Inches(1.45), Inches(4.05), Inches(0.85), Theme.PRIMARY_LIGHT)
    add_text_box(
        slide, Inches(8.75), Inches(1.52), Inches(4.05), Inches(0.32),
        "社内ツール　Step 4と5の間", size=Pt(12), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER,
    )
    add_text_box(
        slide, Inches(8.75), Inches(1.82), Inches(4.05), Inches(0.40),
        "PatentSQUARE", size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER,
    )
    add_text_box(
        slide, Inches(9.0), Inches(2.50), Inches(3.55), Inches(3.95),
        "クレーム案が先行特許と衝突しないかを、本格的に見るときに使う。\n\n"
        "使い方はこのマニュアルでは説明しない。\n\n"
        "参考：知的財産部ホームページ（利用案内）\n"
        f"{PS_REF}",
        size=Pt(13), color=Theme.TEXT,
    )
    _footer(slide, page)
    return slide


def add_template_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "最終の章立て（ANAQUA）", "これは書き順ではない。最後に開示書へ落とすときの形です")
    items = [
        ("1", "技術分野"),
        ("2.1", "従来の図"),
        ("2.2", "従来の問題"),
        ("3.1", "目的"),
        ("3.2", "本発明の図"),
        ("3.3", "具体的内容"),
        ("3.4", "効果"),
        ("4", "発展例"),
        ("5", "クレーム案"),
    ]
    for i, (n, t) in enumerate(items):
        x = Inches(0.45) + Inches(1.4) * i
        add_rect(slide, x, Inches(2.15), Inches(1.28), Inches(2.85), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(0.34), Inches(2.38), Inches(0.55), NAVY)
        add_text_box(slide, x + Inches(0.34), Inches(2.45), Inches(0.55), Inches(0.42), n, size=Pt(11), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, x + Inches(0.08), Inches(3.10), Inches(1.12), Inches(1.6), t, size=Pt(14), color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(0.45), Inches(5.25), Inches(12.4), Inches(1.55), CREAM, corner=True)
    add_text_box(
        slide, Inches(0.65), Inches(5.38), Inches(12.0), Inches(1.30),
        "作業は次のページの Step 0〜5 の順で進める。\n"
        "この章立てに中身を写すのは、いちばん最後（Step 5）です。\n"
        "テンプレート：知的財産部ホームページ（ANAQUA形式）",
        size=Pt(15), color=NAVY, bold=True,
    )
    _footer(slide, page)
    return slide


def add_process_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "本マニュアルの作業順", "上段がやること。下段がゴール。5枚目の章立てとは別物です")
    steps = [
        ("0", "タネを言語化", "課題と解決を\n30秒で話せる"),
        ("1", "先行特許で\n背景と方向性", "近い先行特許と\n方向性が見える"),
        ("2", "本発明を具体化", "構成と流れと\n差が言える"),
        ("3", "差を表にする", "従来との差が\n1枚の表になる"),
        ("4", "クレームと発展", "被りにくい\n守り方が書ける"),
        ("5", "開示書へ落とす", "章立てが埋まり\n矛盾がない"),
    ]
    for i, (n, t, g) in enumerate(steps):
        x = Inches(0.38) + Inches(2.16) * i
        add_rect(slide, x, Inches(1.38), Inches(2.06), Inches(4.55), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(0.72), Inches(1.50), Inches(0.48), NAVY)
        add_text_box(
            slide, x + Inches(0.72), Inches(1.54), Inches(0.48), Inches(0.40), n,
            size=Pt(15), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, x + Inches(0.08), Inches(2.08), Inches(1.9), Inches(0.85), t,
            size=Pt(13), color=Theme.TEXT_MUTED, align=PP_ALIGN.CENTER,
        )
        add_rect(slide, x + Inches(0.12), Inches(2.98), Inches(1.82), Inches(0.05), NAVY)
        add_rect(slide, x + Inches(0.10), Inches(3.15), Inches(1.86), Inches(2.55), GOAL_BG, corner=True)
        add_text_box(
            slide, x + Inches(0.12), Inches(3.22), Inches(1.82), Inches(0.32), "ゴール",
            size=Pt(12), color=GOLD, bold=True, align=PP_ALIGN.CENTER,
        )
        add_text_box(
            slide, x + Inches(0.14), Inches(3.55), Inches(1.78), Inches(2.0), g,
            size=Pt(14), color=NAVY, bold=True, align=PP_ALIGN.CENTER,
        )
        if i < 5:
            add_text_box(
                slide, x + Inches(1.88), Inches(1.55), Inches(0.32), Inches(0.4), "→",
                size=Pt(16), color=GOLD, bold=True,
            )
    add_rect(slide, Inches(0.38), Inches(6.05), Inches(12.55), Inches(1.00), CREAM, corner=True)
    add_text_box(
        slide, Inches(0.55), Inches(6.14), Inches(12.2), Inches(0.82),
        "先行特許は2回見る　Step 1＝背景と方向性　／　Step 4と5の間＝クレーム案の衝突確認（PatentSQUARE）\n"
        "衝突したら Step 2 へ戻る。次のページで、この2回の役割を分けて示す。",
        size=Pt(14), color=NAVY, bold=True,
    )
    _footer(slide, page)
    return slide


def add_two_searches_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "先行特許は2回見る", "同じ調査を繰り返すのではなく、見るものが違う")
    cols = [
        (
            "1回目　Step 1",
            "背景の調査。近い先行特許に被らないよう、クレーム案のおおまかな方向性を決める。",
            "・Copilotで近い特許の見当をつける\n"
            "・似ている点／違う点を一言にする\n"
            "・独立項の核になりそうな方向だけ決める\n"
            "・この段階ではクレームの文言は書かない\n"
            "・PatentSQUAREはまだ使わない",
        ),
        (
            "2回目　Step 4と5の間",
            "作ったクレーム案が、先行特許と衝突するかを判断する。",
            "・ここがPatentSQUAREの出番\n"
            "・使い方はこのマニュアルでは説明しない\n"
            "・参考：知的財産部ホームページ\n"
            f"・{PS_REF}\n"
            "・衝突するなら Step 2 へ戻る\n"
            "・衝突しなければ Step 5 へ進む",
        ),
    ]
    for i, (t, lead, body) in enumerate(cols):
        x = Inches(0.45) + Inches(6.4) * i
        add_rect(slide, x, Inches(1.40), Inches(6.15), Inches(5.40), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.40), Inches(6.15), Inches(0.70), NAVY)
        add_text_box(slide, x, Inches(1.50), Inches(6.15), Inches(0.50), t, size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, x + Inches(0.28), Inches(2.25), Inches(5.60), Inches(1.15), lead, size=Pt(14), color=NAVY, bold=True)
        add_text_box(slide, x + Inches(0.28), Inches(3.50), Inches(5.60), Inches(3.00), body, size=Pt(14), color=Theme.TEXT)
    _footer(slide, page)
    return slide


def add_step_slide(prs, *, page, step, title, goal, check, do_items, question, answer):
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, Theme.SLIDE_WIDTH, Inches(0.10), NAVY)
    add_text_box(
        slide, Inches(0.45), Inches(0.22), Inches(7.5), Inches(0.55),
        f"Step {step}　{title}", size=Pt(22), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    _progress(slide, int(step))

    add_rect(slide, Inches(0.40), Inches(0.88), Inches(12.52), Inches(0.95), NAVY, corner=True)
    add_text_box(slide, Inches(0.60), Inches(0.92), Inches(3.2), Inches(0.28), "このStepのゴール", size=Pt(12), color=GOLD, bold=True)
    add_text_box(slide, Inches(0.60), Inches(1.18), Inches(12.12), Inches(0.55), goal, size=Pt(20), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    add_rect(slide, Inches(0.40), Inches(1.92), Inches(12.52), Inches(0.42), CREAM, corner=True)
    add_text_box(
        slide, Inches(0.55), Inches(1.96), Inches(12.22), Inches(0.34),
        f"次へ進む前の確認　□ {check}", size=Pt(13), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )

    chip_w = Inches(4.08)
    for i, item in enumerate(do_items[:3]):
        x = Inches(0.40) + (chip_w + Inches(0.14)) * i
        add_rect(slide, x, Inches(2.42), chip_w, Inches(0.48), Theme.CARD, corner=True)
        add_text_box(
            slide, x + Inches(0.12), Inches(2.46), chip_w - Inches(0.20), Inches(0.40),
            f"やること {i + 1}　{item}", size=Pt(11), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )

    _qa(slide, Inches(0.40), Inches(3.00), Inches(12.52), Inches(4.05), question, answer)
    _footer(slide, page)
    return slide


def add_polished_slide(prs, *, page, step, title, body):
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, Theme.SLIDE_WIDTH, Inches(0.10), NAVY)
    add_text_box(
        slide, Inches(0.45), Inches(0.22), Inches(7.5), Inches(0.55),
        f"Step {step}　仕上げ（③）　{title}", size=Pt(20), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    _progress(slide, int(step))
    add_rect(slide, Inches(0.40), Inches(0.88), Inches(12.52), Inches(0.58), ANS, corner=True)
    add_text_box(
        slide, Inches(0.60), Inches(0.96), Inches(12.12), Inches(0.42),
        "3枚目の③　人と議論したあと、Copilotでもう一度整えた版。①より量が多く、開示書に近い。",
        size=Pt(14), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_rect(slide, Inches(0.40), Inches(1.58), Inches(12.52), Inches(5.40), Theme.CARD, corner=True)
    add_rect(slide, Inches(0.40), Inches(1.58), Inches(0.12), Inches(5.40), ANS)
    _paras(
        slide, Inches(0.70), Inches(1.72), Inches(11.95), Inches(5.10),
        body, size=Pt(12), color=Theme.TEXT, space_after=Pt(4),
    )
    _footer(slide, page)
    return slide


def add_conflict_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "Step 4と5の間　クレーム案の衝突確認", "ここがPatentSQUAREの出番。衝突したら Step 2 へ戻る")
    add_rect(slide, Inches(0.45), Inches(1.40), Inches(12.4), Inches(1.15), NAVY, corner=True)
    add_text_box(slide, Inches(0.65), Inches(1.48), Inches(3.2), Inches(0.28), "この確認のゴール", size=Pt(12), color=GOLD, bold=True)
    add_text_box(
        slide, Inches(0.65), Inches(1.78), Inches(12.0), Inches(0.60),
        "作ったクレーム案が、先行特許と衝突しないと判断できる",
        size=Pt(18), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    boxes = [
        ("やること", "PatentSQUAREで、独立項の核に近い文献を見る。使い方は本マニュアルでは説明しない。"),
        ("判断", "衝突する → Step 2へ戻り、構成と差をずらす。\n衝突しない → Step 5へ進み、ANAQUAの章立てへ落とす。"),
        ("参考", f"知的財産部ホームページ（PatentSQUARE利用案内）\n{PS_REF}"),
    ]
    for i, (t, b) in enumerate(boxes):
        y = Inches(2.75) + Inches(1.35) * i
        add_rect(slide, Inches(0.45), y, Inches(12.4), Inches(1.22), Theme.CARD, corner=True)
        add_text_box(slide, Inches(0.70), y + Inches(0.12), Inches(2.2), Inches(0.98), t, size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, Inches(3.00), y + Inches(0.18), Inches(9.55), Inches(0.90), b, size=Pt(14), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE)
    _footer(slide, page)
    return slide


def add_graphviz_policy_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "図はCopilotの中だけで作る", "社外のGraphvizサイトやオンラインツールは使わない")
    cards = [
        ("なぜ社外ツールを使わないか", "発明のタネを外部サイトへ貼ると、情報漏れになる。Graphviz Online などは使わない。"),
        ("何をCopilotに頼むか", "① GraphvizのDOTを書く　② 同じ会話で、その図を出す。外部サービスは指定しない。"),
        ("自分が見ること", "素子名・矢印・分岐が、開示書の文章と一致するか。図がきれいでも中身が違うなら直す。"),
    ]
    for i, (t, b) in enumerate(cards):
        y = Inches(1.40) + Inches(1.75) * i
        add_rect(slide, Inches(0.5), y, Inches(12.3), Inches(1.60), Theme.CARD, corner=True)
        add_circle(slide, Inches(0.75), y + Inches(0.45), Inches(0.70), NAVY)
        add_text_box(slide, Inches(0.75), y + Inches(0.55), Inches(0.70), Inches(0.50), str(i + 1), size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, Inches(1.70), y + Inches(0.22), Inches(10.8), Inches(0.45), t, size=Pt(18), color=NAVY, bold=True)
        add_text_box(slide, Inches(1.70), y + Inches(0.75), Inches(10.8), Inches(0.65), b, size=Pt(15), color=Theme.TEXT)
    _footer(slide, page)
    return slide


def add_graphviz_howto_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "Copilotへの聞き方（Graphviz）", "2回に分けて頼む。例は安全計装ロジック自動構築ツール")
    add_rect(slide, Inches(0.45), Inches(1.40), Inches(12.4), Inches(2.35), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(1.50), Inches(12.0), Inches(0.35), "1回目　DOTを書かせる", size=Pt(16), color=ASK, bold=True)
    _paras(
        slide, Inches(0.65), Inches(1.90), Inches(12.0), Inches(1.70),
        "『安全計装ロジック自動構築ツールの従来フローを、GraphvizのDOTで書いてください。"
        "HAZOP→手作業のC&E→手作業のFBD→別ツールのSIL検証、の順です。"
        "社外のレンダリングサービスは使わず、コードだけ出してください。』",
        size=Pt(14), color=Theme.TEXT, space_after=Pt(2),
    )
    add_rect(slide, Inches(0.45), Inches(3.90), Inches(12.4), Inches(2.90), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(4.00), Inches(12.0), Inches(0.35), "2回目　同じ会話で図を出させる", size=Pt(16), color=ANS, bold=True)
    _paras(
        slide, Inches(0.65), Inches(4.42), Inches(12.0), Inches(2.20),
        "『今のDOTを、この会話の中でフローチャートの図にしてください。外部サイトへ貼らないでください。"
        "図が出ないときは、PowerPointの図形で同じ流れを書いてください。』\n"
        "出てきた図を見て、名前と矢印が文章と合うか確認する。合わなければ『検証から合成へ戻る矢印を足して』と直させる。",
        size=Pt(14), color=Theme.TEXT, space_after=Pt(3),
    )
    _footer(slide, page)
    return slide


def add_graphviz_example_slide(prs, *, page, title, subtitle, prompt, nodes, note, extra_nodes=None, box_w=None, gap=None):
    slide = blank_slide(prs)
    header_bar(slide, title, subtitle)
    add_rect(slide, Inches(0.45), Inches(1.35), Inches(12.4), Inches(1.35), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(1.42), Inches(12.0), Inches(0.28), "Copilotへの依頼", size=Pt(13), color=ASK, bold=True)
    add_text_box(slide, Inches(0.65), Inches(1.72), Inches(12.0), Inches(0.85), prompt, size=Pt(13), color=Theme.TEXT)
    add_text_box(slide, Inches(0.45), Inches(2.80), Inches(12.4), Inches(0.32), "Copilotが出した図の例（このマニュアルの安全計装の例）", size=Pt(13), color=ANS, bold=True)
    _flow_row(slide, nodes, Inches(3.20), left=Inches(0.45), box_w=box_w or Inches(2.20), box_h=Inches(0.80), gap=gap or Inches(0.38))
    if extra_nodes:
        add_text_box(slide, Inches(0.45), Inches(4.15), Inches(12.4), Inches(0.28), extra_nodes[0], size=Pt(12), color=Theme.TEXT_MUTED)
        _flow_row(slide, extra_nodes[1], Inches(4.48), left=Inches(0.45), box_w=box_w or Inches(2.20), box_h=Inches(0.80), gap=gap or Inches(0.38))
    add_text_box(slide, Inches(0.45), Inches(5.55 if extra_nodes else 4.20), Inches(12.4), Inches(1.40), note, size=Pt(14), color=NAVY)
    _footer(slide, page)
    return slide


def build_deck():
    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT
    n = 1

    def pg():
        nonlocal n
        n += 1
        return str(n)

    add_title_slide(
        prs,
        title="Copilotで作る\n発明開示書",
        subtitle="ゼロから素案まで　─　作業順と、最終の章立ては別",
        footer="王 者興  |  MKDS BDD Gr1.  |  2026/06/12",
    )
    add_purpose_slide(prs, pg())
    add_guardrail_slide(prs, pg())
    add_tools_slide(prs, pg())
    add_template_slide(prs, pg())
    add_process_slide(prs, pg())
    add_two_searches_slide(prs, pg())

    add_step_slide(
        prs,
        page=pg(),
        step="0",
        title="発明のタネを言葉にする",
        goal="課題と解決の骨格を、口頭で30秒話せる",
        check="背景・困っていること・何をしたいか、を自分の言葉で言える",
        do_items=[
            "タイトルと「何が困るか」を聞く",
            "メモや会議動画があれば入れる",
            "穴のある文章でよいので出す",
        ],
        question="『安全計装ロジック自動構築ツール』という発明アイデアがあります。プラント安全計装設計における課題と、本発明の概要を整理してください。権利範囲（クレーム）はまだ作らないでください。",
        answer=(
            "了解です。いまは開示書の「背景・課題・本発明・効果」の骨格だけ作ります。数字や工数は仮なので、自部署の実績に置き換えてください。\n"
            "■ 背景　安全計装（SIS）の設計では、HAZOPで危険シナリオを出し、原因と結果の表（C&E）とファンクションブロック図（FBD）を人手で作り、別ツールでSIL検証することが多いです。入力はHAZOP表・P&ID・機器リスト、出力はSIFのロジックと検証記録です。担当者がツールをまたいで転記し、作図しています。\n"
            "■ 課題　①転記と作図に時間がかかる　②担当者でロジックの切り方・命名がぶれる　③HAZOP項目の漏れや、設計変更が検証に届かない　④「なぜこの素子が入ったか」が成果物に残らない。\n"
            "■ 本発明　HAZOPなどの入力から、知識ベースと規則でSIFロジックを自動合成し、SIL検証までつなぎ、C&E／FBDと追跡表として出すツールです。人が見る点（例外、特殊インターロック）は残します。\n"
            "■ 効果　工数削減、属人化の低減、漏れ低減、変更時の影響範囲の可視化（トレーサビリティ）。\n"
            "この骨格で30秒話せるようになったら、次は近い先行特許を探します。『自動生成』という言葉だけで権利の話に入らないでください。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="0",
        title="発明のタネ",
        body=(
            "技術分野：プラントの安全計装（SIS）設計。対象は、HAZOPの結果からSIF（安全計装機能）のロジックを組み立て、SIL検証までつなぐ作業である。\n"
            "背景：現状の設計では、HAZOPワークショップで危険シナリオを出したあと、原因と結果の表（C&E）とファンクションブロック図（FBD）を担当者が手作業で作る。SIL検証は別ツールで行い、結果を設計書へ転記する。入力はHAZOP表、P&ID、機器リスト。出力はSIFロジックと検証記録である。ツールが分かれているため、同じ情報を何度も写している。\n"
            "従来の問題：転記と作図に時間がかかる。担当者ごとにロジックの切り方と命名がぶれる。HAZOP項目の漏れや、設計変更が検証側に届かないことがある。「なぜこの素子が入ったか」が成果物に残らず、レビューと変更影響の把握が属人的になる。\n"
            "目的：HAZOP等の入力からSIFロジックを合成し、検証し、成果物と追跡表として出すことで、工数・属人化・漏れ・変更追跡の問題を同時に減らす。\n"
            "本発明の概要：入力部、知識ベース、合成エンジン、検証部、出力部を持つ。人が見る点（例外シナリオ、特殊インターロック、不合格時の方針）は残す。自動生成だけで終わらせない。\n"
            "効果：工数削減、切り方のそろえ、漏れの早期発見、変更時の影響範囲の可視化。数字は自部署の実績に置き換える。\n"
            "次に進む条件：この骨格を、資料を見ずに30秒で話せる。"
        ),
    )

    add_step_slide(
        prs,
        page=pg(),
        step="1",
        title="先行特許で背景と方向性を見る",
        goal="近い先行特許が見え、クレーム案のおおまかな方向性が決まる",
        check="被りそうな点を避けた方向を、一言で言える",
        do_items=[
            "Copilotで近い先行特許を聞く",
            "似ている点／違う点を整理する",
            "方向だけ決める（文言は書かない）",
        ],
        question="本発明に近い先行特許を探すキーワードと、調査で見る観点を出してください。権利範囲（クレーム）はまだ作らないでください。ヒットした『自動生成』だけで本発明と同じと思わないようにしてください。",
        answer=(
            "了解です。クレームの文言は書きません。先に『何が近いか』を見る検索語と、読む観点だけ出します。まずはCopilotで近い特許の見当をつけてください。本格的な衝突確認は、クレーム案を作ったあと（Step 4と5の間）にPatentSQUAREで行います。\n"
            "■ 検索の核　安全計装、SIS、SIF、HAZOP、C&E、cause and effect、FBD、インターロック、トリップロジック、ロジック自動生成、SIL検証、トレーサビリティ。英語も併用：safety instrumented system, automatic generation, cause and effect matrix。\n"
            "■ 組み合わせ例　「HAZOP 自動 AND SIF」「cause and effect 自動生成 安全計装」「SIL verification 自動 ロジック」。広すぎたら『検証』『追跡』『HAZOP』を必須語にして絞る。\n"
            "■ 読む観点　①ロジックをどこまで自動で作るか　②SIL検証とつながるか　③HAZOP項目と成果物の対応が残るか　④人がどこを確認するか　⑤対象がSISか、DCSの一般ロジックか。\n"
            "■ 方向性の決め方　『自動生成』は先行特許にも多い。方向の核は『検証結果を設計へ戻す』『入力項目と素子の対応を残す』に置く。文言はまだ書かない。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="1",
        title="背景調査と方向性",
        body=(
            "調査の目的（1回目）：背景をつかむことと、あとで書くクレーム案が近い先行特許に正面から被らないよう、方向だけ決めること。衝突の最終判断はしない。PatentSQUAREは使わない。\n"
            "検索語（確定版）：安全計装／SIS／SIF／HAZOP／C&E／FBD／インターロック／ロジック自動生成／SIL検証／トレーサビリティ。英語：safety instrumented system, cause and effect matrix, automatic generation, SIL verification。必須の絞り込み語は「検証」または「追跡」。\n"
            "近い先行特許の型（Copilotが見立てた例）：A ルールでC&Eを自動生成する　B インターロックロジックをライブラリから組み立てる　C SIL計算を支援する　D HAZOP支援だが成果物への追跡がない。いずれも『生成まで』または『検証だけ』で止まっていることが多い。\n"
            "似ている点：HAZOPやルールからロジックを自動で作る、という言い方。違う点：検証結果を設計成果物へ戻すか、HAZOP項目と素子の対応表を出すか、人が例外を確認する点を残すか。\n"
            "決めた方向（クレームの核の候補）：独立項の核は『自動生成』に置かない。『合成＋検証＋設計への反映＋対応の記録』に置く。文言はStep 4まで書かない。\n"
            "次に進む条件：似ている点／違う点と、避ける方向を一言で言える。"
        ),
    )

    add_step_slide(
        prs,
        page=pg(),
        step="2",
        title="本発明を具体化する",
        goal="構成・処理の流れ・先行特許にない点が説明できる",
        check="入力→処理→出力と、「先行特許にないところ」を指差せる",
        do_items=[
            "Step 1で見た先行特許を踏まえる",
            "入力・処理・出力に分けて書く",
            "先行特許にないところだけ厚くする",
        ],
        question="先行特許にはルールベースのロジック自動生成があります。本発明の違いを、システム構成と処理の流れで具体化してください。人が確認する点も残してください。",
        answer=(
            "先行特許の『ルールでロジックを自動生成する』は残しつつ、本発明が厚いところ（検証と追跡）を構成と流れで書きます。ここが後の差の表とクレームの材料になります。\n"
            "■ 構成　①入力部：HAZOP表、機器リスト、必要ならP&IDのタグ　②知識ベース：SIFパターン、禁止組み合わせ、命名規則　③合成エンジン：シナリオからC&E／FBDを生成　④検証部：SIL計算、入力とロジックの整合チェック　⑤出力部：設計書、追跡表　⑥人が見る点：例外シナリオ、特殊インターロック、不合格時の方針。\n"
            "■ 流れ　取り込む → シナリオをSIF候補にする → 知識ベースを参照してロジックを合成する → 検証する → 成果物と追跡表を出す。不合格なら合成に戻る。人が確認してから確定する。\n"
            "■ 先行特許との差　先行特許は生成までが多い。本発明は検証結果を設計成果物へ戻し、HAZOPの原因−SIF−素子の対応を残す。自動生成そのものは差になりにくいので、ここを具体例つきで書いてください。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="2",
        title="構成と流れ",
        body=(
            "システム構成（開示書3.3向け）：(1)入力部はHAZOP表、機器リスト、任意でP&IDタグを取り込む。(2)知識ベースはSIFパターン、禁止組み合わせ、命名規則を持つ。(3)合成エンジンはシナリオからC&EとFBDを生成する。(4)検証部はSIL計算と、入力項目とロジックの整合を見る。(5)出力部は設計書と追跡表を出す。(6)人が見る点は、例外シナリオ、特殊インターロック、不合格時の方針である。\n"
            "処理の流れ：取り込む → シナリオをSIF候補にする → 知識ベースを参照して合成する → 検証する → 成果物と追跡表を出す。不合格なら合成に戻る。人が確認してから確定する。確定後に設計変更が入った場合は、追跡表から影響範囲を出して再検証する。\n"
            "先行特許にないところ（厚く書く）：検証結果を設計成果物へ戻すこと。HAZOPの原因−SIF−素子の対応を記録すること。不合格時に合成へ戻る閉じたループを持つこと。人が例外を確認する点を、システム構成の一部として残すこと。\n"
            "先行特許と重なりやすいところ（薄く書く）：ルールベースであること、ロジックを自動生成すること自体。ここを独立項の核にしない。\n"
            "図にする対象：図1 従来フロー　図2 本発明の構成　図3 処理の流れ　図4 追跡表の例。図はあとでCopilotにGraphvizで作らせる。社外ツールは使わない。"
        ),
    )

    add_step_slide(
        prs,
        page=pg(),
        step="3",
        title="従来技術との差を表にする",
        goal="従来／課題／差／優位性が、1枚の表で対比できる",
        check="表の「差」の行を、人が見ても同じ理解になる",
        do_items=[
            "先行特許の要約を入力する",
            "本発明との差だけ残す",
            "差がない項目は狙いをずらす",
        ],
        question="先行特許は「ルールでロジックを自動生成するが、SIL検証と変更追跡はない」とします。本発明との差を、従来／課題／差／優位性の表にしてください。差が弱い欄も指摘してください。",
        answer=(
            "仮定どおり『生成はあるが、検証と追跡はない』として対比します。表は開示書の2.2（従来の問題）と3.4（効果）にそのまま使えます。\n"
            "■ 従来　HAZOP後にC&E／FBDを手作業。一部ツールはルールでロジックだけ自動生成する。検証と設計変更の反映は別作業のまま。\n"
            "■ 従来の課題　転記ミスと属人化が残る。生成結果の根拠が残らない。変更が入ると、どのHAZOP項目とどの素子を直すかが追えない。\n"
            "■ 差　本発明は合成に加え、検証部を持ち、検証結果を成果物へ戻す。HAZOP原因−SIF−素子の追跡表を出す。不合格時は合成に戻る。人が例外を確認する点を残す。\n"
            "■ 優位性　漏れに気づきやすい。レビューが追跡表でできる。変更時に影響範囲が見える。担当者による切り方のぶれを知識ベースで抑えやすい。\n"
            "■ 差が弱い欄　『自動生成すること自体』『ルールベースであること』は先行特許と重なりやすい。クレームの核にしない。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="3",
        title="差の表",
        body=(
            "対比表（開示書2.2／3.4に写す）：\n"
            "項目／従来／本発明\n"
            "ロジック作成　手作業、またはルールで生成まで　合成したあと検証し、不合格なら合成に戻る\n"
            "検証　別ツール。結果の転記は手作業　検証部を持ち、結果を設計成果物へ戻す\n"
            "変更追跡　担当者の記憶と個別メモ　HAZOP原因−SIF−素子の追跡表を出す\n"
            "人の関与　切り方も確認も人に依存　例外と特殊インターロックだけ人が確認する\n"
            "差の核：検証結果の設計への反映、入力項目と素子の対応の記録。差が弱い欄：自動生成、ルールベース。核にしない。\n"
            "優位性：漏れに気づきやすい。レビューが追跡表でできる。変更時に影響範囲が見える。命名と切り方のぶれを知識ベースで抑えやすい。\n"
            "使い方：この表を人が読んで、差の行の意味が同じになるか確認する。同じにならなければ、Step 2の構成の書き方を直してから表を更新する。"
        ),
    )

    add_step_slide(
        prs,
        page=pg(),
        step="4",
        title="クレーム案と発展例",
        goal="先行特許と被りにくい保護範囲の案がある",
        check="独立項が「差の核」だけになっており、自動生成だけでは終わっていない",
        do_items=[
            "独立項は差の核だけにする",
            "従属項で変形・適用を広げる",
            "将来の派生テーマも一行で残す",
        ],
        question="上の差を踏まえ、先行特許と被りにくい独立クレームの骨格と、従属クレーム・発展例の案を作ってください。独立項が『自動生成』だけで終わらないようにしてください。",
        answer=(
            "独立項は差の核だけにします。『ロジックを自動生成する手段』だけで始めると、Step 1で見た先行特許に被りやすいです。\n"
            "■ 独立項の核　HAZOP等の入力を取り込む手段と、SIFロジックを合成する手段と、合成結果を検証する手段と、検証結果を設計成果物へ反映する手段と、入力項目とロジック素子の対応を記録する手段、を備える。ポイントは『検証して戻す』と『対応を記録する』が落ちないこと。\n"
            "■ 従属の例　知識ベースを更新する、不合格時に再合成する、追跡表を出力する、人が例外を確認してから確定する、複数プラントへ展開する、P&IDタグと対応づける。\n"
            "■ 発展例　運転中のインターロック見直し、異常検知ロジックへの展開、規則の学習による知識ベース更新。\n"
            "■ 派生テーマ　検証結果の説明文生成、HAZOP支援ツールとの接続、変更差分だけの再検証。\n"
            "次はPatentSQUAREで、この骨格が先行特許と衝突しないかを見る。衝突したらStep 2へ戻る。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="4",
        title="クレーム案",
        body=(
            "独立項の骨格（差の核だけ）：HAZOP等の危険シナリオを取り込む手段、知識ベースを参照してSIFロジックを合成する手段、合成結果を検証する手段、検証結果を設計成果物へ反映する手段、入力項目とロジック素子の対応を記録する手段、を備えるシステム。自動生成する手段、だけでは独立項にしない。\n"
            "従属項：知識ベースを更新する／不合格時に再合成する／追跡表を出力する／人が例外を確認してから確定する／P&IDタグと対応づける／複数プラントへ展開する。\n"
            "発展例：運転中のインターロック見直し。異常検知ロジックへの展開。規則の学習による知識ベース更新。\n"
            "派生テーマ（一行）：検証結果の説明文生成。HAZOP支援ツールとの接続。変更差分だけの再検証。\n"
            "まだやらないこと：この文言をANAQUAに登録すること。先に、Step 4と5の間で衝突確認をする。\n"
            "次に進む条件：独立項を読んで、『自動生成だけ』に見えない。見えたら核を差の2点（戻す／記録する）に戻す。"
        ),
    )
    add_conflict_slide(prs, pg())

    add_step_slide(
        prs,
        page=pg(),
        step="5",
        title="ANAQUAの項目へ落とす",
        goal="各項目が埋まり、課題・手段・効果・クレームが矛盾していない",
        check="課題で書いた困りごとが、手段・効果・クレームのどこかに対応している",
        do_items=[
            "公式フォーマットに写す",
            "用語をSIFなどにそろえる",
            "矛盾チェックだけAIに頼む",
        ],
        question="課題・手段・効果・クレームに矛盾がないか見てください。「安全計装ロジック」「SIF」「インターロック」の用語もそろえてください。抜けている対応があれば指摘してください。",
        answer=(
            "並べて見ます。よくあるずれは『課題は漏れと属人化なのに、クレームが自動生成だけ』です。効果の『漏れ防止』が権利の核に残っていない、という状態です。\n"
            "■ 対応の確認　2.2の課題（転記・属人化・漏れ・変更が届かない）→ 3.3の構成（合成＋検証＋追跡）→ 3.4の効果（工数・漏れ・影響範囲）→ 5の独立項（検証して戻す、対応を記録する）。どれかが『自動生成』だけに戻っていたら、差の核が落ちています。\n"
            "■ 用語　本文とクレームで『安全計装ロジック』と『SIF』を混在させない。初出で『SIF（安全計装機能）』と定義し、以降はSIFにそろえる。インターロックはSIFを実現するロジックの呼び方として書く。\n"
            "■ 仕上げ　公式フォーマットへ写したあと、課題の困りごとが手段・効果・クレームのどこかに矢印でつながるかだけ見る。つながらない行は削るか、構成側を厚くする。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="5",
        title="章立てへ落とす",
        body=(
            "写す先（5枚目の章立て）：1 技術分野＝プラントSIS設計。2.1 従来の図＝手作業のC&E／FBD。2.2 従来の問題＝転記・属人化・漏れ・変更が届かない。3.1 目的＝合成と検証と追跡を一つの流れにする。3.2 本発明の図＝入力・知識・合成・検証・出力。3.3 具体的内容＝Step 2の仕上げ。3.4 効果＝工数・漏れ・影響範囲。4 発展例＝Step 4の発展。5 クレーム案＝衝突確認を通った独立項と従属項。\n"
            "用語（そろえた定義）：SIF（安全計装機能）。初出で定義し、以降はSIF。インターロックはSIFを実現するロジック。C&Eは原因と結果の表。FBDはファンクションブロック図。SISは安全計装システム。\n"
            "矛盾チェック結果（例）：課題の『漏れと変更が届かない』は、手段の検証部と追跡表、効果の漏れ低減と影響範囲、独立項の『戻す／記録する』に対応している。『自動生成』だけに戻っている行はない。\n"
            "まだやらないこと：Copilotの出力をそのままANAQUAへ登録すること。人が説明できる文だけ残す。"
        ),
    )

    add_graphviz_policy_slide(prs, pg())
    add_graphviz_howto_slide(prs, pg())
    add_graphviz_example_slide(
        prs,
        page=pg(),
        title="図1　従来フロー",
        subtitle="手作業のC&E／FBDと、別ツールのSIL検証",
        prompt="『従来のSIS設計フローをGraphvizのDOTで書いて。HAZOP→手作業C&E→手作業FBD→別ツールSIL検証。社外サイトは使わず、この会話で図も出して。』",
        nodes=["HAZOP", "手作業 C&E", "手作業 FBD", "別ツール\nSIL検証"],
        note="見るところ：矢印がすべて『手作業／別ツール』になっているか。本発明の検証ループが混ざっていないか。",
        box_w=Inches(2.20),
        gap=Inches(0.38),
    )
    add_graphviz_example_slide(
        prs,
        page=pg(),
        title="図2　本発明の構成",
        subtitle="入力・知識・合成・検証・出力。人が見る点も残す",
        prompt="『本発明の構成をGraphvizのDOTで書いて。入力部、知識ベース、合成エンジン、検証部、出力部、人が見る点。社外サイトは使わず、この会話で図も出して。』",
        nodes=["入力部", "知識ベース", "合成エンジン", "検証部", "出力部"],
        extra_nodes=("人が見る点は、検証のあとに残す。", ["例外シナリオ", "特殊\nインターロック", "不合格時の方針"]),
        note="見るところ：6つの箱がそろっているか。人が見る点が消えて『全自動』になっていないか。",
        box_w=Inches(2.20),
        gap=Inches(0.38),
    )
    add_graphviz_example_slide(
        prs,
        page=pg(),
        title="図3　処理の流れ",
        subtitle="不合格なら合成に戻る。人が確認してから確定する",
        prompt="『処理の流れをGraphvizのDOTで書いて。取り込み→SIF候補→合成→検証→出力。不合格なら合成に戻る。人が確認して確定。社外サイトは使わず、この会話で図も出して。』",
        nodes=["取り込む", "SIF候補", "合成する", "検証する", "出力する"],
        extra_nodes=("不合格のとき", ["検証する", "合成に戻る", "人が確認", "確定する"]),
        note="見るところ：検証から合成へ戻る矢印があるか。人が確認する箱が抜けていないか。",
        box_w=Inches(2.20),
        gap=Inches(0.38),
    )
    add_graphviz_example_slide(
        prs,
        page=pg(),
        title="図4　トレーサビリティの例",
        subtitle="HAZOPの原因 − SIF − 素子、の対応が残ること",
        prompt="『追跡の例をGraphvizのDOTで書いて。HAZOP原因→SIF→ロジック素子。変更が入ったら影響範囲が見えるように。社外サイトは使わず、この会話で図も出して。』",
        nodes=["HAZOP原因", "SIF", "ロジック素子", "影響範囲"],
        note="見るところ：3つが一本の線でつながっているか。変更から影響範囲へ矢印があるか。図が開示書3.3の文章と同じか。",
        box_w=Inches(2.20),
        gap=Inches(0.38),
    )
    return prs
