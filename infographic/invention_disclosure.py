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
        slide, Inches(0.45), Inches(0.22), Inches(7.6), Inches(0.55),
        f"Step {step}　仕上げ（③）　{title}", size=Pt(18), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    _progress(slide, int(step))
    add_rect(slide, Inches(0.40), Inches(0.82), Inches(12.52), Inches(0.48), ANS, corner=True)
    add_text_box(
        slide, Inches(0.60), Inches(0.88), Inches(12.12), Inches(0.36),
        "3枚目の③。人と直したあとCopilotで整えた、ANAQUAに入れる分量の文例。",
        size=Pt(13), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_rect(slide, Inches(0.40), Inches(1.40), Inches(12.52), Inches(5.60), Theme.CARD, corner=True)
    add_rect(slide, Inches(0.40), Inches(1.40), Inches(0.12), Inches(5.60), ANS)
    _paras(
        slide, Inches(0.65), Inches(1.50), Inches(12.05), Inches(5.38),
        body, size=Pt(11), color=Theme.TEXT, space_after=Pt(3),
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
    header_bar(slide, "Copilotへの聞き方（Graphviz）", "2回に分けて頼む。例はプロセス開発高速化ツール")
    add_rect(slide, Inches(0.45), Inches(1.40), Inches(12.4), Inches(2.35), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(1.50), Inches(12.0), Inches(0.35), "1回目　DOTを書かせる", size=Pt(16), color=ASK, bold=True)
    _paras(
        slide, Inches(0.65), Inches(1.90), Inches(12.0), Inches(1.70),
        "『プロセス開発高速化ツールの処理フローを、GraphvizのDOTで書いてください。"
        "入力受付→候補実験条件群の生成→プロセスモデルによるシミュレーション→目的関数による評価→"
        "有望条件がなければ条件再設定、あれば次回実験条件の選定→条件と選定理由の出力、です。"
        "社外のレンダリングサービスは使わず、コードだけ出してください。』",
        size=Pt(14), color=Theme.TEXT, space_after=Pt(2),
    )
    add_rect(slide, Inches(0.45), Inches(3.90), Inches(12.4), Inches(2.90), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(4.00), Inches(12.0), Inches(0.35), "2回目　同じ会話で図を出させる", size=Pt(16), color=ANS, bold=True)
    _paras(
        slide, Inches(0.65), Inches(4.42), Inches(12.0), Inches(2.20),
        "『今のDOTを、この会話の中でフローチャートの図にしてください。外部サイトへ貼らないでください。"
        "図が出ないときは、PowerPointの図形で同じ流れを書いてください。』\n"
        "出てきた図を見て、名前と矢印が文章と合うか確認する。合わなければ『有望条件がないとき、候補生成へ戻る矢印を足して』と直させる。",
        size=Pt(14), color=Theme.TEXT, space_after=Pt(3),
    )
    _footer(slide, page)
    return slide


def _prompt_box(slide, prompt):
    add_rect(slide, Inches(6.85), Inches(1.35), Inches(6.00), Inches(2.35), Theme.CARD, corner=True)
    add_text_box(slide, Inches(7.05), Inches(1.42), Inches(5.60), Inches(0.28), "Copilotへの依頼", size=Pt(13), color=ASK, bold=True)
    add_text_box(slide, Inches(7.05), Inches(1.74), Inches(5.60), Inches(1.80), prompt, size=Pt(12), color=Theme.TEXT)


def _note_box(slide, note):
    add_rect(slide, Inches(6.85), Inches(3.85), Inches(6.00), Inches(2.95), CREAM, corner=True)
    add_text_box(slide, Inches(7.05), Inches(3.95), Inches(5.60), Inches(0.32), "見るところ", size=Pt(13), color=GOLD, bold=True)
    add_text_box(slide, Inches(7.05), Inches(4.32), Inches(5.60), Inches(2.30), note, size=Pt(13), color=NAVY)


def _vbox(slide, x, y, w, h, text, *, fill=NAVY, color=None):
    add_rect(slide, x, y, w, h, fill, corner=True)
    add_text_box(
        slide, x + Inches(0.06), y + Inches(0.04), w - Inches(0.12), h - Inches(0.08),
        text, size=Pt(11), color=color or Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


def add_gv_conventional_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "図1　従来フロー", "実験とシミュレーションを併用し、装置設計を最適化する従来技術")
    x, w = Inches(0.70), Inches(5.50)
    items = [
        (Inches(1.32), Inches(0.36), "スタート", GOLD),
        (Inches(1.74), Inches(0.62), "S101 実験等による実験  ／  S102 数学的モデル化・数値シミュレーション", NAVY),
        (Inches(2.42), Inches(0.52), "S103 最適化処理", NAVY),
        (Inches(3.00), Inches(0.58), "S104 目的関数は収束したか？　N→S102へ戻る", Theme.PRIMARY_LIGHT),
        (Inches(3.64), Inches(0.52), "S105 実験等による実験", NAVY),
        (Inches(4.22), Inches(0.58), "S106 制約条件を逸脱していないか？　N→S103へ戻る", Theme.PRIMARY_LIGHT),
        (Inches(4.86), Inches(0.36), "終了", GOLD),
    ]
    for y, h, t, fill in items:
        _vbox(slide, x, y, w, h, t, fill=fill)
    _prompt_box(
        slide,
        "『特開2003-281194に代表される従来技術のフローをGraphvizのDOTで書いて。"
        "実験と数値シミュレーションを最適化処理へ入れ、目的関数が収束するまで戻し、"
        "収束後に実験し、制約逸脱なら最適化へ戻る。社外サイトは使わず、この会話で図も出して。』",
    )
    _note_box(
        slide,
        "・実験（S101／S105）がループの中にあるか。\n"
        "・主眼が装置設計の最適化になっており、次回実験条件を事前に提示する箱がないか。\n"
        "・本発明の「候補を実験前に高速評価する」流れが混ざっていないか。",
    )
    _footer(slide, page)
    return slide


def add_gv_architecture_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "図2　システム構成図", "入力部・演算部・出力部。3.3の構成と対応させる")
    add_text_box(slide, Inches(0.45), Inches(1.22), Inches(12.4), Inches(0.28), "Copilotが出した図の例（プロセス開発高速化ツール）", size=Pt(12), color=ANS, bold=True)
    cols = [
        ("入力部", ["目的関数\n（評価指標）", "過去実験データ\n（実測値・条件）", "プロセスモデル\n（装置構成・運転条件）"]),
        ("演算部", ["候補条件生成部\n（条件探索空間の設定）", "シミュレーション実行部\n（高速・網羅的評価）", "評価・選定部\n（目的関数スコアリング）"]),
        ("出力部", ["次回実験条件\n（推奨パラメータ一式）", "選定理由・根拠\n（シミュレーション結果）", "レポート\n（評価情報の紐付け）"]),
    ]
    for i, (head, boxes) in enumerate(cols):
        x = Inches(0.45) + Inches(4.25) * i
        add_rect(slide, x, Inches(1.58), Inches(4.08), Inches(3.55), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.58), Inches(4.08), Inches(0.42), NAVY)
        add_text_box(slide, x, Inches(1.62), Inches(4.08), Inches(0.34), head, size=Pt(14), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        for j, t in enumerate(boxes):
            _vbox(slide, x + Inches(0.18), Inches(2.14) + Inches(0.92) * j, Inches(3.72), Inches(0.82), t, fill=GOAL_BG, color=NAVY)
    add_rect(slide, Inches(0.45), Inches(5.28), Inches(12.4), Inches(1.60), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.65), Inches(5.38), Inches(12.0), Inches(0.28), "Copilotへの依頼", size=Pt(13), color=ASK, bold=True)
    add_text_box(
        slide, Inches(0.65), Inches(5.70), Inches(12.0), Inches(1.05),
        "『プロセス開発高速化ツールのシステム構成をGraphvizのDOTで書いて。入力部（目的関数、過去実験データ、プロセスモデル）、"
        "演算部（候補条件生成、シミュレーション実行、評価・選定）、出力部（次回実験条件、選定理由）。社外サイトは使わず、この会話で図も出して。』\n"
        "見るところ：3.3の入力部・演算部・出力部と箱が一致しているか。選定理由の箱が落ちていないか。",
        size=Pt(12), color=Theme.TEXT,
    )
    _footer(slide, page)
    return slide


def add_gv_process_slide(prs, page):
    slide = blank_slide(prs)
    header_bar(slide, "図3　処理フローチャート", "有望条件がなければ候補生成へ戻る")
    x, w = Inches(0.55), Inches(6.10)
    steps = [
        (Inches(1.32), Inches(0.34), "開始", GOLD),
        (Inches(1.70), Inches(0.50), "S10 入力受付（目的関数／過去実験データ／プロセスモデル）", NAVY),
        (Inches(2.24), Inches(0.50), "S20 候補実験条件群の生成（探索空間の設定・サンプリング）", NAVY),
        (Inches(2.78), Inches(0.50), "S30 プロセスモデルによるシミュレーション（推定結果を高速算出）", NAVY),
        (Inches(3.32), Inches(0.50), "S40 目的関数による評価（候補条件のスコアリング）", NAVY),
        (Inches(3.86), Inches(0.50), "有望条件が存在するか？　NO→S20へ戻る（条件再設定）", Theme.PRIMARY_LIGHT),
        (Inches(4.40), Inches(0.50), "S50 次回実験条件の選定（上位スコア案件の絞り込み）", NAVY),
        (Inches(4.94), Inches(0.50), "S60 実験条件と選定理由の出力（推奨パラメータ／根拠）", NAVY),
        (Inches(5.48), Inches(0.34), "終了", GOLD),
    ]
    for y, h, t, fill in steps:
        _vbox(slide, x, y, w, h, t, fill=fill)
    _prompt_box(
        slide,
        "『処理フローをGraphvizのDOTで書いて。入力受付→候補生成→シミュレーション→目的関数評価。"
        "有望条件がなければ候補生成へ戻る。あれば次回実験条件を選定し、条件と選定理由を出力。"
        "社外サイトは使わず、この会話で図も出して。』",
    )
    _note_box(
        slide,
        "・S20へ戻る矢印（条件再設定）があるか。\n"
        "・実験そのものはフローの外で、計算ループが先に回っているか。\n"
        "・出力に「選定理由」が残っているか。条件だけになっていないか。",
    )
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
        question="『プロセス開発高速化ツール』という発明アイデアがあります。プロセス開発における実験計画の課題と、本発明の概要を整理してください。権利範囲（クレーム）はまだ作らないでください。",
        answer=(
            "了解です。いまは開示書の「背景・課題・本発明・効果」の骨格だけ作ります。数字は自部署の実績に置き換えてください。\n"
            "■ 背景　化学プロセスや材料開発では、収率や品質を上げるために、原料組成・量・温度などの流入条件と、温度・pHなどの処理条件を変えながら実験を繰り返します。1回の実験に数時間〜数日かかることがあります。\n"
            "■ 課題　①実験計画が担当者の経験に依存する　②実験の前に、多数の候補条件を高速に評価する仕組みが弱い　③実験は最適化結果の検証として後段に置かれ、無駄な実験を事前に削れない　④なぜその条件を選んだかの説明が残りにくい。\n"
            "■ 本発明　目的関数、過去の実験データ、装置情報に基づくプロセスモデルを入力し、シミュレーションで次回の有望な実験条件と選定理由を出すツールです。\n"
            "■ 効果　実験回数の削減、開発リードタイムの短縮、条件選定の属人性の低減、検討過程の説明性の向上。\n"
            "この骨格で30秒話せるようになったら、次は近い先行特許を探します。『シミュレーションで最適化する』だけで権利の話に入らないでください。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="0",
        title="1.技術分野／3.1目的",
        body=(
            "1. 技術分野\n"
            "本発明は、化学プロセス、材料開発、反応・分離プロセス等におけるプロセス開発分野に適用され、実験計画、条件探索および開発効率化を支援するシミュレーション活用技術に関する。本発明が他に応用できる可能性として、スケールアップ検討、条件変更時の事前検討、既存プロセスの改良検討、バイオプロセス開発などが挙げられる。\n"
            "2.2 従来技術の問題（骨格）\n"
            "従来、収率や品質の向上に向けた条件検討は、原料組成・量・温度等の流入条件および温度・pH等の処理条件を変更しながら、数時間〜数日単位で実験を反復する必要があった。特開2003-281194等に代表される従来技術は、実験結果と数値シミュレーション結果を併用して装置設計パラメータを最適化するが、次回実験条件を実験前に体系的に選定・提示する仕組みは十分に示されていない。そのため、実験計画段階での意思決定の属人性や、実験回数削減による開発リードタイム短縮という課題が残されている。\n"
            "3.1 本発明の目的\n"
            "本発明の目的は、プロセス開発における実験計画および条件探索を効率化・高速化することである。目的関数、過去の実験データ、ならびに装置情報に基づくプロセスモデルを活用し、シミュレーションにより次回実験の有望条件を事前に導出することで、経験依存の判断や無駄な試行を削減し、開発期間およびコストの低減を図る。"
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
        question="本発明に近い先行特許を探すキーワードと、調査で見る観点を出してください。権利範囲（クレーム）はまだ作らないでください。ヒットした『シミュレーションで最適化』だけで本発明と同じと思わないようにしてください。",
        answer=(
            "了解です。クレームの文言は書きません。近い特許の見当と、読む観点だけ出します。本格的な衝突確認は、クレーム案のあと（Step 4と5の間）にPatentSQUAREで行います。\n"
            "■ 検索の核　プロセスモデル、実験計画、条件探索、シミュレーション、次回実験条件、目的関数。分野：モデリング、モデル開発、プロセスモデル。\n"
            "■ 近い文献の型　A 実験とシミュレーションを併用して装置設計を最適化する（特開2003-281194）　B 過去データでモデルを調整し制御を良くする　C 探索空間の次の実験条件を決める情報処理。いずれも装置設計・運転制御・データ駆動探索が主眼になりやすい。\n"
            "■ 方向　独立項の核は『装置パラメータの最適化』に置かない。『実験前に候補条件をプロセスモデルで評価し、次回実験条件と選定理由を出す』に置く。文言はまだ書かない。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="1",
        title="出願前調査（検索と文献）",
        body=(
            "【出願前調査の結果】\n"
            "■AI検索の検索条件\n"
            "（分野（IPC分類））モデリング、モデル開発、プロセスモデル\n"
            "（検索自然文）本発明は、プロセス開発における実験計画・条件探索を高速化する「プロセス開発高速化ツール」に関する。従来、収率や品質の向上に向けた条件検討は、原料組成・量・温度等の流入条件および温度・pH等の処理条件を変更しながら、数時間〜数日単位で実験を反復する必要があった。本ツールは、装置情報（構成、形状、サイズ等）から構築したモデル（動的モデル、プロセスモデル、物理モデル）を用いて実験結果をシミュレーションにより推定し、次回の有望な実験条件を導出することで、実験回数と検討リードタイムを低減する。\n"
            "（AIスコア（高/中/低））中\n"
            "＜先行文献＞特開2003-281194、特許第3771858号、特開2013-069094、特開2025-139210、特開2025-152843\n"
            "＜従来技術の構成／動作＞特開2003-281194等に代表される従来技術では、装置に関する実機による実験結果と数値シミュレーション結果とを併用し、最適化手法およびデータ分析手法を用いて装置設計を最適化する最適設計支援技術を開示している。実験データ解析装置、数値シミュレーション装置、および最適設計支援装置を連携させ、実験では取得困難な物理量とシミュレーションでは得られない実測値とを補完的に用いることで、装置パラメータの最適化を図る。さらに、最適化結果を用いた再実験を繰り返すことで、効率的かつ高精度な装置設計を実現することを目的としている。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="1",
        title="先行文献の評価",
        body=(
            "＜先行文献の評価＞\n"
            "先行文献群には、実験データや数値シミュレーションを用いて装置設計や運転条件を最適化する技術、あるいは探索効率を向上させる技術が複数開示されている。例えば、実験結果とシミュレーション結果を併用して装置パラメータを最適化する技術や、過去データを用いてモデルを調整し制御性能を向上させる技術、さらには探索空間における次の実験条件を決定する情報処理技術などが提案されている。これらはいずれも、従来の試行錯誤型の検討に比べ、効率化を図ろうとする点に特徴を有している。\n"
            "しかしながら、これらの先行文献においては、主たる目的が装置設計の最適化、運転制御の高度化、あるいはデータ駆動型探索に置かれており、プロセス開発段階における実験計画そのものを体系的に支援する構成は必ずしも十分に開示されていない。特に、実験実施前の段階で、目的関数および過去の実験データを踏まえつつ、プロセスモデルを用いて多数の候補条件を高速に評価し、その中から「次回実施すべき実験条件」を合理的根拠とともに提示する点については、明確な記載が見られない。\n"
            "また、先行文献の多くは、最適化結果や探索結果を数値的に導出することに主眼が置かれており、なぜその条件が選定されたのかという説明性や、検討過程の再現性については十分に考慮されていない。そのため、条件選定の判断が依然として担当者の経験や解釈に依存しやすく、開発現場における意思決定の属人性を完全には解消できないという課題が残されている。\n"
            "これに対し、本発明は、目的関数、過去の実験データおよび装置情報に基づくプロセスモデルを入力として、実験前にシミュレーションを用いた高速な条件評価を行い、次回実験条件とその選定理由を一体的に出力する点に特徴を有する。この構成により、無駄な実験の削減、開発リードタイムの短縮のみならず、検討プロセスの可視化および説明性の向上が可能となる点で、先行文献とは技術的思想および効果の面で明確に相違している。\n"
            "決めた方向：独立項の核は装置設計の最適化に置かない。実験前評価と、次回条件＋選定理由の一体出力に置く。"
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
        question="先行特許には実験とシミュレーションを併用した装置設計の最適化があります。本発明の違いを、入力部・演算部・出力部と処理の流れで具体化してください。",
        answer=(
            "先行特許の『実験とシミュレーションを併用する』は残しつつ、本発明が厚いところ（実験前の候補評価と、次回条件＋選定理由）を構成で書きます。\n"
            "■ 構成　入力部：目的関数、過去の実験データ、プロセスモデル。演算部：候補実験条件の生成、プロセスモデルによるシミュレーション、目的関数に基づく評価と次回条件の選定。出力部：選定した実験条件と、選定理由となるシミュレーション結果・評価情報。\n"
            "■ 流れ　入力を受け付ける → 候補条件群を生成する → 各候補をシミュレーションする → 目的関数で評価する → 有望条件がなければ候補生成へ戻る → 次回実験条件を選定する → 条件と選定理由を出力する。\n"
            "■ 先行特許との差　先行は装置設計パラメータの最適化が主眼で、実験は最適化結果の検証として後段に置かれる。本発明は実験実施前に多数の候補を高速評価し、次回実施すべき実験条件を根拠つきで提示する。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="2",
        title="3.3 構成／動作",
        body=(
            "＜本発明具体例の構成／動作＞\n"
            "本発明の具体例に係るプロセス開発高速化ツールは、入力部、演算部および出力部を備えて構成される。入力部には、プロセス開発における目的関数、過去の実験データ、ならびに装置構成や運転条件に基づいて構築されたプロセスモデルが入力される。\n"
            "演算部は、前記入力情報に基づき、複数の候補となる実験条件を生成し、各候補条件についてプロセスモデルを用いたシミュレーションを実行することで、対応する推定結果を算出する。さらに、算出された推定結果を目的関数の観点から評価し、次回実施すべき有望な実験条件を選定する。\n"
            "出力部は、選定された実験条件とともに、当該条件が選定された理由となるシミュレーション結果や評価情報を関連付けて出力する。これにより、実験実施前に結果の見込みと判断根拠を把握したうえで、効率的に実験を進めることが可能となる。\n"
            "処理の流れ（図3）：S10 入力受付（目的関数／過去実験データ／プロセスモデル）→ S20 候補実験条件群の生成（条件探索空間の設定・サンプリング）→ S30 プロセスモデルによるシミュレーション（各候補条件について推定結果を高速算出）→ S40 目的関数による評価（各候補条件の推定結果をスコアリング）→ 有望条件が存在するか。存在しなければS20へ戻り条件を再設定する。存在すれば S50 次回実験条件の選定（上位スコア案件の絞り込み）→ S60 実験条件と選定理由の出力（推奨パラメータ／根拠シミュレーション結果）。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="2",
        title="3.3 効果",
        body=(
            "＜本発明具体例の効果＞\n"
            "本発明によれば、プロセス開発において従来必要とされていた「条件変更―実験―結果解釈―次条件設定」という時間集約的な実験反復ループに先立ち、数秒から数分で完了する計算ループを高速に実行することが可能となる。これにより、実験実施前の段階で、多数の候補条件をシミュレーションにより網羅的に評価でき、情報価値の低い条件を事前に排除したうえで、目的関数に対して寄与度の高い実験条件を選定できる。この結果、実験回数の削減および開発リードタイムの短縮が実現され、開発コストの低減に大きく寄与する。\n"
            "また、本発明では、目的関数、過去の実験データ、ならびに装置情報に基づき構築されたプロセスモデルを統合的に用いるため、単なる経験則や局所的な試行に依存することなく、条件探索を論理的かつ再現性のある形で実行できる。特に、候補条件ごとに想定される実験結果を事前に算出し、その根拠となるシミュレーション情報とともに提示することで、次回実験条件の選定理由が明確化され、検討過程の説明性および客観性が向上する。この作用により、条件選定の属人性が低減され、担当者間や組織間での知見共有や引き継ぎが容易となる。\n"
            "さらに、プロセスモデルを活用することで、実験では直接計測が困難な内部状態量や時間変化挙動を推定することが可能となり、実験結果の解釈精度が向上する。これにより、単に最適条件を探索するだけでなく、プロセス挙動や支配因子に対する理解が深化し、開発初期段階から合理的な条件設定や制約整理が可能となる。その結果、後戻りの少ない効率的なプロセス開発が実現され、複雑な制約条件を有するプロセスにおいても、安定した開発推進を可能とするという効果を奏する。"
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
        question="先行特許は「実験とシミュレーションを併用して装置設計を最適化するが、次回実験条件を実験前に提示しない」とします。本発明との差を、従来／課題／差／優位性の表にしてください。",
        answer=(
            "仮定どおり『併用はあるが、実験計画そのものは支援しない』として対比します。表は開示書の2.2と3.3の効果に使えます。\n"
            "■ 従来　実験結果と数値シミュレーションを併用し、装置パラメータを最適化する。実験は最適化結果の検証として後段に置かれる。\n"
            "■ 従来の課題　次回実験条件を事前に体系的に選定・提示する仕組みが弱い。実験計画が属人的。無駄な実験を実験前に削りにくい。選定理由が残りにくい。\n"
            "■ 差　本発明は実験実施前に、目的関数・過去データ・プロセスモデルで多数の候補を高速評価し、次回実験条件と選定理由を一体で出す。\n"
            "■ 差が弱い欄　『シミュレーションを使うこと』『最適化すること』は先行と重なりやすい。核にしない。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="3",
        title="2.2 問題点と差",
        body=(
            "＜従来技術の問題点＞\n"
            "先行特許に示される従来技術では、実験結果とシミュレーション結果を併用する点では本発明と共通するものの、主眼は装置設計パラメータの最適化にあり、次回実験条件を事前に体系的に選定・提示する仕組みは十分に示されていない。また、実験は最適化結果の検証として位置付けられており、実験前に多数の候補条件を高速に評価して無駄な実験を削減するという発想には至っていない。そのため、実験計画段階での意思決定の属人性や、実験回数削減による開発リードタイム短縮という課題は依然として残されている。\n"
            "対比（ANAQUA 2.2／3.3に写す）\n"
            "項目／従来／本発明\n"
            "主眼　装置設計パラメータの最適化　プロセス開発における実験計画・条件探索の支援\n"
            "実験の位置　最適化結果の検証として後段　実験実施前に候補を評価し、実施すべき条件を先に出す\n"
            "出力　最適化された装置パラメータ　次回実験条件と、選定理由となるシミュレーション結果・評価情報\n"
            "説明性　数値的な最適解の導出が主　なぜその条件かを根拠つきで提示し、検討過程を再現できる\n"
            "差の核：実験前の高速な候補評価、次回実験条件と選定理由の一体出力。差が弱い欄：シミュレーションの利用、最適化そのもの。独立項の核にしない。"
        ),
    )

    add_step_slide(
        prs,
        page=pg(),
        step="4",
        title="クレーム案と発展例",
        goal="先行特許と被りにくい保護範囲の案がある",
        check="独立項が「差の核」だけになっており、装置最適化だけでは終わっていない",
        do_items=[
            "独立項は差の核だけにする",
            "従属項で変形・適用を広げる",
            "将来の派生テーマも一行で残す",
        ],
        question="上の差を踏まえ、先行特許と被りにくい独立クレームの骨格と、従属クレーム・発展例の案を作ってください。独立項が『シミュレーションで最適化する』だけで終わらないようにしてください。",
        answer=(
            "独立項は差の核だけにします。『シミュレーションで装置を最適化する手段』だけだと、特開2003-281194に被りやすいです。\n"
            "■ 独立項の核　目的関数、過去の実験データ、装置構成や運転条件に基づき生成した複数の候補実験条件をプロセスモデルに入力する入力部と、各候補についてシミュレーションを実行し推定結果を算出する演算部、を備える装置または方法。\n"
            "■ 従属　推定結果を目的関数で評価し次回実験条件を選定する／選定した条件を入力部へ戻す（フィードバック）／プロセスモデルの構築部／選定理由を関連付けて出力する出力部（レポート）。\n"
            "■ 発展　物理モデル／機械学習モデル／ハイブリッド、複数の目的関数、化学以外（材料、分離、バイオ）、クラウド分散。\n"
            "次はPatentSQUAREで衝突を見る。衝突したらStep 2へ戻る。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="4",
        title="3.5 要点／3.4 発展",
        body=(
            "権利化を希望する本発明の要点\n"
            "・独立項案　プロセス開発における目的関数、過去の実験データ、ならびに装置構成や運転条件（入力情報）に基づき生成した複数の候補となる実験条件をプロセスモデルに入力する入力部、各候補実験条件についてプロセスモデルを用いたシミュレーションを実行することで、対応する推定結果を算出する演算部、から構成された装置、方法。\n"
            "・従属項案１　算出された推定結果を目的関数に基づいて評価し、次回実施すべき（有望な）実験条件を選定する。\n"
            "・従属項案２　選定した実験条件を入力部に入力する（フィードバックループ化）。\n"
            "・従属項案３　前記入力情報に基づいてプロセスモデルを構築する構築部を備えたもの。\n"
            "・従属項案４　選定された実験条件とともに、当該条件が選定された理由となるシミュレーション結果や評価情報を関連付けて出力する出力部を備えたもの（レポート機能）。\n"
            "3.4 応用／変形／水平展開\n"
            "本発明におけるプロセスモデルは、物理モデル、機械学習モデル、またはこれらを組み合わせたハイブリッドモデルであってもよく、対象とするプロセスや開発段階に応じて適宜選択・切替されてもよい。動的モデル／静的モデルは、物理モデル／経験モデル（機械学習モデル）とは別次元として選択できる。目的関数は、収率や品質指標の最大化に限定されず、コスト低減、エネルギー消費最小化、制約条件の満足度向上など、複数の評価指標を組み合わせて設定することが可能である。さらに、化学プロセスに限らず、材料開発、反応条件最適化、分離・精製プロセス設計、バイオプロセス開発などへ水平展開でき、研究開発段階のみならずスケールアップ検討、条件変更時の事前検討、既存プロセスの改良検討にも適用可能である。実装は単独の計算機に限らず、クラウド環境やネットワークを介した分散処理としてもよい。"
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
            "用語をプロセスモデルにそろえる",
            "矛盾チェックだけAIに頼む",
        ],
        question="課題・手段・効果・クレームに矛盾がないか見てください。「プロセスモデル」「シミュレーション」「最適化」の用語もそろえてください。抜けている対応があれば指摘してください。",
        answer=(
            "並べて見ます。よくあるずれは『課題は実験計画の属人性なのに、クレームがシミュレーションによる最適化だけ』です。効果の『次回条件を根拠つきで出す』が権利の核に残っていない、という状態です。\n"
            "■ 対応の確認　2.2の課題（実験計画が属人、無駄な実験を事前に削れない、選定理由が残らない）→ 3.3の構成（入力・演算・出力、実験前評価）→ 3.3の効果（回数削減、説明性）→ 独立項（候補をプロセスモデルへ入力し推定する）と従属（選定、理由の出力）。どれかが『最適化する』だけに戻っていたら、差の核が落ちています。\n"
            "■ 用語　動的モデル／物理モデル／プロセスモデルを混在させない。本文では「プロセスモデル」にそろえ、必要なら「動的かつ物理的なプロセスモデル」と定義する。機械学習モデルは経験モデルとして、動的／静的とは別次元で書く。"
        ),
    )
    add_polished_slide(
        prs,
        page=pg(),
        step="5",
        title="章立てへ落とす",
        body=(
            "写す先（ANAQUA形式）\n"
            "1 技術分野＝化学プロセス、材料開発、反応・分離プロセス等のプロセス開発。実験計画・条件探索を支援するシミュレーション活用技術。\n"
            "2.1 従来の図＝実験と数値シミュレーションを最適化処理へ入れ、収束するまで戻し、その後に実験するフロー（図1）。\n"
            "2.2 従来の問題＝主眼が装置設計パラメータの最適化。次回実験条件の事前提示が弱い。実験計画の属人性、リードタイム。先行文献の評価をここに添付する。\n"
            "3.1 目的＝実験計画および条件探索の効率化・高速化。次回実験の有望条件を事前に導出する。\n"
            "3.2 本発明の図＝図2 システム構成図（入力部・演算部・出力部）、図3 処理フローチャート（S10〜S60）。\n"
            "3.3 具体的内容＝入力部・演算部・出力部の構成／動作。3.3 効果＝計算ループを実験ループの前に回す、選定理由の提示、内部状態の推定。\n"
            "4 発展例＝モデル種、目的関数、水平展開、クラウド。\n"
            "5 クレーム案＝衝突確認を通った独立項（入力部＋演算部）と従属項1〜4。\n"
            "用語：本文とクレームで「プロセスモデル」にそろえる。矛盾チェック：課題の属人性・無駄な実験は、手段の実験前評価と理由つき出力、独立項・従属項4に対応している。『最適化する』だけに戻っている行はない。出力をそのまま登録しない。"
        ),
    )

    add_graphviz_policy_slide(prs, pg())
    add_graphviz_howto_slide(prs, pg())
    add_gv_conventional_slide(prs, pg())
    add_gv_architecture_slide(prs, pg())
    add_gv_process_slide(prs, pg())
    return prs
