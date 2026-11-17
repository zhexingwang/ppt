"""Copilotとのやり取りのコツ。発明開示書マニュアルの補足。"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .shapes import add_circle, add_rect, add_text_box, header_bar, set_paragraph_text
from .slides import add_title_slide, blank_slide
from .theme import Theme


NAVY = Theme.PRIMARY
BLUE = Theme.PRIMARY_LIGHT
RED = RGBColor(0xC0, 0x39, 0x2B)
GREEN = Theme.ACCENT_2
CREAM = RGBColor(0xFD, 0xF6, 0xE3)
PALE = RGBColor(0xE8, 0xF1, 0xFA)
MINT = RGBColor(0xE7, 0xF4, 0xEA)
ROSE = RGBColor(0xFD, 0xF2, 0xF0)
ZEBRA = RGBColor(0xF3, 0xF7, 0xFB)

LEFT = Inches(0.42)
CONTENT_W = Inches(12.50)


def _footer(slide, page: str):
    add_text_box(
        slide, Inches(0.45), Inches(7.18), Inches(8.5), Inches(0.26),
        "Confidential(Yokogawa)", size=Pt(10), color=Theme.TEXT_MUTED,
    )
    add_text_box(
        slide, Inches(11.6), Inches(7.18), Inches(1.2), Inches(0.26),
        page, size=Pt(10), color=Theme.TEXT_MUTED, align=PP_ALIGN.RIGHT,
    )


def _paras(slide, left, top, width, height, text: str, *, size=Pt(14), color=Theme.TEXT, bold=False, space_after=Pt(2)):
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


def _new(prs, title: str, subtitle: str, page: str):
    slide = blank_slide(prs)
    header_bar(slide, title, subtitle)
    _footer(slide, page)
    return slide


def _card(slide, x, y, w, h, fill=Theme.CARD):
    add_rect(slide, x, y, w, h, fill, corner=True)
    return (x, y, w, h)


def _num(slide, x, y, n: str, fill=NAVY, size=Inches(0.34)):
    add_circle(slide, x, y, size, fill)
    add_text_box(
        slide, x, y, size, size, n,
        size=Pt(12), color=Theme.WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


def add_opening_slide(prs, page):
    slide = _new(prs, "作業を分けないと、文章が混ざる", "一つの特許を書いているときの、チャットの使い方です", page)
    problems = [
        (
            "同じチャットで作業を続けている",
            "本文、先行特許の整理、クレーム、図、見直しを、\n一つのチャットのまま進めています。\n前の作業の指示が残ると、今の文章に混ざります。",
        ),
        (
            "依頼が短すぎる",
            "「いい感じに書いて」だけだと、渡した事実が少ない分、前の作業の内容で空いたところを補ってしまいます。",
        ),
    ]
    for i, (title, body) in enumerate(problems):
        x = LEFT + (Inches(6.35)) * i
        _card(slide, x, Inches(1.28), Inches(6.15), Inches(2.55))
        add_rect(slide, x, Inches(1.28), Inches(6.15), Inches(0.58), NAVY)
        add_text_box(
            slide, x + Inches(0.22), Inches(1.36), Inches(5.7), Inches(0.42),
            title, size=Pt(18), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        _paras(
            slide, x + Inches(0.24), Inches(2.04), Inches(5.7), Inches(1.55),
            body, size=Pt(15),
        )
    add_text_box(
        slide, LEFT, Inches(4.02), CONTENT_W, Inches(0.32),
        "このリストの中身", size=Pt(14), color=NAVY, bold=True,
    )
    items = [
        ("1", "作業ごとに分ける", "本文、先行特許、クレーム、\n図と見直しは、チャットを分けます。"),
        ("2", "最初に書く", "今の作業、成果物、決まった事実。\n守ることと、出力の形も先に書きます。"),
        ("3", "直し方を指定する", "残す文と、変える箇所を指定します。一度に頼むことは一つです。"),
    ]
    for i, (n, title, body) in enumerate(items):
        x = LEFT + Inches(4.20) * i
        _card(slide, x, Inches(4.40), Inches(4.02), Inches(2.45))
        _num(slide, x + Inches(0.20), Inches(4.58), n)
        add_text_box(
            slide, x + Inches(0.64), Inches(4.58), Inches(3.15), Inches(0.36),
            title, size=Pt(18), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        _paras(slide, x + Inches(0.22), Inches(5.15), Inches(3.58), Inches(1.45), body, size=Pt(14))
    return slide


def add_why_slide(prs, page):
    slide = _new(prs, "同じチャットは、前の作業を前提に答える", "一つの特許の中でも、作業が変わると前提が残ります", page)
    cards = [
        (
            NAVY,
            "残るのは、このチャットのやり取り",
            "入力した文章は、モデルの学習には使いません。今開いているチャットのやり取りが、次の回答の前提になります。",
        ),
        (
            BLUE,
            "前の作業の指示が残る",
            "最初からのやり取りを読んだうえで、次の回答を作ります。前の作業の指示が残っていると、今の依頼より、前の指示の影響が大きくなることがあります。",
        ),
    ]
    for i, (color, title, body) in enumerate(cards):
        x = LEFT + Inches(6.35) * i
        _card(slide, x, Inches(1.28), Inches(6.15), Inches(2.85))
        add_rect(slide, x, Inches(1.28), Inches(0.10), Inches(2.85), color)
        add_text_box(
            slide, x + Inches(0.28), Inches(1.42), Inches(5.7), Inches(0.42),
            title, size=Pt(18), color=color, bold=True,
        )
        _paras(slide, x + Inches(0.28), Inches(1.95), Inches(5.65), Inches(1.95), body, size=Pt(14))

    add_text_box(
        slide, LEFT, Inches(4.28), CONTENT_W, Inches(0.30),
        "混ざったときに見えること", size=Pt(14), color=NAVY, bold=True,
    )
    symptoms = [
        ("先行特許の言い回しが残る", "文献整理の言葉が、\n本発明の説明に入ります。"),
        ("やめた構成が戻る", "採用しなかった手段が、\nクレームに再び入ります。"),
        ("図と本文の符号がずれる", "図で検討した要素が、\n本文の構成と合いません。"),
    ]
    for i, (title, body) in enumerate(symptoms):
        x = LEFT + Inches(4.20) * i
        _card(slide, x, Inches(4.64), Inches(4.02), Inches(1.28), CREAM)
        add_text_box(
            slide, x + Inches(0.18), Inches(4.74), Inches(3.66), Inches(0.40),
            title, size=Pt(15), color=NAVY, bold=True,
        )
        add_text_box(
            slide, x + Inches(0.18), Inches(5.16), Inches(3.66), Inches(0.60),
            body, size=Pt(13), color=Theme.TEXT,
        )
    add_text_box(
        slide, LEFT, Inches(6.08), CONTENT_W, Inches(0.90),
        "作業が変わったら、新しいチャットを開きます。\n「これまでの指示は忘れて」と書いても、同じチャットでは前の指示が残ることがあります。",
        size=Pt(15), color=NAVY, bold=True,
    )
    return slide


def add_when_slide(prs, page):
    slide = _new(prs, "作業が変わったら、チャットを新しくする", "同じチャットで続けるのは、同じ成果物を直しているあいだだけです", page)
    columns = [
        (NAVY, Theme.WHITE, "新しく開く", [
            ("作業の種類が変わった", "本文の次に、クレーム、図、見直しへ進むときは分けます。"),
            ("従来の図と本発明の図", "図の種類が違うときは、チャットも分けます。"),
            ("書いた文章を見直したい", "書いたチャットとは別にし、決まった原稿だけを貼ります。"),
            ("方針を変えて書き直す", "やめた構成が残ったままだと、また文章に入ります。"),
            ("指示が消えないとき", "そのチャットは閉じ、決まったことだけを次に書きます。"),
        ]),
        (GREEN, Theme.WHITE, "同じチャットで続ける", [
            ("同じ箇所を直している", "同じ段落や、同じ請求項の続きだけを直します。"),
            ("直前の箇所を指している", "「この表の2行目だけ」のように、直前の出力を指します。"),
            ("一つの図を仕上げている", "図の指示を書かせた直後は、同じチャットで図にします。"),
        ]),
    ]
    widths = [Inches(6.55), Inches(5.75)]
    xs = [LEFT, LEFT + Inches(6.75)]
    for (header, header_color, label, rows), x, w in zip(columns, xs, widths):
        add_rect(slide, x, Inches(1.28), w, Inches(0.48), header, corner=True)
        add_text_box(
            slide, x, Inches(1.32), w, Inches(0.40),
            label, size=Pt(16), color=header_color, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        top = Inches(1.86)
        row_h = Inches(0.78) if len(rows) == 5 else Inches(1.30)
        for i, (title, body) in enumerate(rows):
            y = top + row_h * i
            _card(slide, x, y, w, row_h - Inches(0.10))
            inset = Inches(0.24) if len(rows) == 3 else Inches(0.08)
            add_text_box(
                slide, x + Inches(0.18), y + inset, w - Inches(0.34), Inches(0.30),
                title, size=Pt(15), color=NAVY, bold=True,
            )
            add_text_box(
                slide, x + Inches(0.18), y + inset + Inches(0.32), w - Inches(0.34), Inches(0.42),
                body, size=Pt(13), color=Theme.TEXT,
            )
    add_text_box(
        slide, LEFT, Inches(6.28), CONTENT_W, Inches(0.62),
        "新しいチャットでも、依頼が短いと、開いている開示書の別の節が混ざることがあります。\n使う材料は、このチャットに貼った文章だけと書きます。",
        size=Pt(14), color=Theme.TEXT,
    )
    return slide


def add_map_slide(prs, page):
    slide = _new(prs, "一つの特許の中で、作業ごとに分ける", "最初の行に、発明の名前と、今の作業を書きます", page)
    headers = [("チャット", Inches(2.15)), ("ここでやること", Inches(5.35)), ("ここではやらないこと", Inches(5.00))]
    x = LEFT
    y = Inches(1.28)
    for text, w in headers:
        add_rect(slide, x, y, w, Inches(0.46), NAVY)
        add_text_box(
            slide, x + Inches(0.10), y, w - Inches(0.16), Inches(0.46),
            text, size=Pt(14), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        x += w
    rows = [
        ("本文", "課題、構成、動作、効果を書く", "図の描き直し、クレームの文言"),
        ("先行特許", "近い文献を整理し、本発明との差を残す", "クレームの清書、効果の文章"),
        ("クレーム", "独立項を決め、そのあと従属項を書く", "背景の長い説明、図の修正"),
        ("図（従来）", "従来の処理の流れ、または従来の構成", "本発明の構成を描き足すこと"),
        ("図（本発明）", "構成図と、処理の流れ", "従来の図の修正"),
        ("見直し", "矛盾、用語のゆれ、抜けを指摘する", "指摘と同時に、全文を書き換える"),
    ]
    row_h = Inches(0.78)
    for i, (a, b, c) in enumerate(rows):
        y = Inches(1.74) + row_h * i
        fill = Theme.CARD if i % 2 == 0 else ZEBRA
        vals = [(a, Inches(2.15), True), (b, Inches(5.35), False), (c, Inches(5.00), False)]
        x = LEFT
        for text, w, strong in vals:
            add_rect(slide, x, y, w, row_h, fill)
            add_text_box(
                slide, x + Inches(0.12), y, w - Inches(0.20), row_h,
                text, size=Pt(14), color=NAVY if strong else Theme.TEXT, bold=strong,
                anchor=MSO_ANCHOR.MIDDLE,
            )
            x += w
    add_text_box(
        slide, LEFT, Inches(6.50), CONTENT_W, Inches(0.55),
        "チャット名の例は「プロセス開発高速化ツール／効果」です。名前は最初の文から付くことが多いので、あとから探せます。",
        size=Pt(14), color=Theme.TEXT,
    )
    return slide


def add_carry_slide(prs, page):
    slide = _new(prs, "次の作業へ渡すもの", "前のチャットの全文は貼りません。決まったことだけを、先頭に書きます", page)
    columns = [
        (GREEN, MINT, "渡すもの", [
            ("発明の一文", "この発明が何をするかを、短く書きます。"),
            ("構成と符号", "決まったものだけを書きます。候補は書きません。"),
            ("用語の統一", "ゆれやすい語を、どの語にそろえるかを書きます。"),
            ("今回書くもの", "どの節か、文体、文字数を書きます。"),
        ]),
        (RED, ROSE, "渡さないもの", [
            ("採用しなかった構成", "やめた案は、次のチャットに書きません。"),
            ("前の作業の指示", "「文献を広く探して」などの指示は残しません。"),
            ("迷っていたやり取り", "結論だけを残し、途中の往復は残しません。"),
            ("図の描き直しの履歴", "採用しなかった要素が、本文に戻る原因になります。"),
        ]),
    ]
    for i, (color, bg, label, rows) in enumerate(columns):
        x = LEFT + Inches(6.35) * i
        add_rect(slide, x, Inches(1.28), Inches(6.15), Inches(0.50), color, corner=True)
        add_text_box(
            slide, x, Inches(1.32), Inches(6.15), Inches(0.42),
            label, size=Pt(16), color=Theme.WHITE, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        for j, (title, body) in enumerate(rows):
            y = Inches(1.90) + Inches(1.05) * j
            _card(slide, x, y, Inches(6.15), Inches(0.96), bg)
            add_text_box(
                slide, x + Inches(0.20), y + Inches(0.08), Inches(5.75), Inches(0.32),
                title, size=Pt(16), color=color, bold=True,
            )
            add_text_box(
                slide, x + Inches(0.20), y + Inches(0.42), Inches(5.75), Inches(0.42),
                body, size=Pt(14), color=Theme.TEXT,
            )
    _card(slide, LEFT, Inches(6.20), CONTENT_W, Inches(0.78), CREAM)
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(6.32), CONTENT_W - Inches(0.40), Inches(0.54),
        "前のチャットをそのまま貼ると、前の作業の指示まで残ります。決まったことだけを、短く書き直して貼ります。",
        size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    return slide


def add_opening_five_slide(prs, page):
    slide = _new(prs, "最初に書く5つのこと", "例は、プロセス開発高速化ツールの効果の段落です", page)
    _card(slide, LEFT, Inches(1.26), CONTENT_W, Inches(4.55), PALE)
    rows = [
        ("作業", "プロセス開発高速化ツールの「効果」の段落だけを書く"),
        ("成果物", "である調。300字から400字。見出しは付けない"),
        ("事実", "1. 入力は、目的関数、過去の実験データ、プロセスモデルである\n2. シミュレーションで、次回実験の有望条件を事前に出す\n3. 有望条件がなければ、候補の生成へ戻る"),
        ("守ること", "数値、短縮の割合、先行特許の番号は作らない。ここにない事実は〔要確認〕と書く\n参照してよいのは、ここに書いた内容だけです"),
        ("出力", "本文だけ。前置きは付けない"),
    ]
    y = Inches(1.40)
    heights = [Inches(0.46), Inches(0.46), Inches(1.20), Inches(0.92), Inches(0.46)]
    for (label, body), h in zip(rows, heights):
        add_text_box(
            slide, LEFT + Inches(0.20), y, Inches(1.45), h,
            label, size=Pt(14), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, LEFT + Inches(1.70), y, Inches(10.50), h,
            body, size=Pt(14), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += h
    _card(slide, LEFT, Inches(5.98), CONTENT_W, Inches(0.95), CREAM)
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(6.08), CONTENT_W - Inches(0.40), Inches(0.32),
        "避けたい頼み方：「いい感じに、特許の文章にして」",
        size=Pt(15), color=NAVY, bold=True,
    )
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(6.42), CONTENT_W - Inches(0.40), Inches(0.40),
        "こう頼むと、前の作業の内容で、書いていないところを埋めてしまいます。",
        size=Pt(14), color=Theme.TEXT,
    )
    return slide


def add_one_request_slide(prs, page):
    slide = _new(prs, "一度に頼むことは、一つだけ", "直すときも、変える箇所は一つにします", page)
    rows = [
        ("1", "作業を同時に頼まない", "書くこと、短くすること、用語をそろえること、見直しは、別の依頼にします。開示書の全体も、1回では頼みません。"),
        ("2", "直す範囲を引用で示す", "直したい段落の名前か、直したい文を引用して範囲を決めます。「全体を見直して」は、見直し用のチャットで使います。"),
        ("3", "残す文を先に示す", "残したい文は「この文は残す」と引用します。指定がないと、その文まで書き換わります。"),
        ("4", "指摘と修正を分ける", "欠点を聞くときは、直す前に問題点だけ出させます。採用するものだけを、次の依頼にします。"),
    ]
    for i, (n, title, body) in enumerate(rows):
        y = Inches(1.24) + Inches(1.08) * i
        _card(slide, LEFT, y, CONTENT_W, Inches(1.00))
        _num(slide, LEFT + Inches(0.16), y + Inches(0.32), n)
        add_text_box(
            slide, LEFT + Inches(0.64), y + Inches(0.06), Inches(11.5), Inches(0.32),
            title, size=Pt(16), color=NAVY, bold=True,
        )
        add_text_box(
            slide, LEFT + Inches(0.64), y + Inches(0.40), Inches(11.55), Inches(0.52),
            body, size=Pt(14), color=Theme.TEXT,
        )
    _card(slide, LEFT, Inches(5.66), CONTENT_W, Inches(1.22), CREAM)
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(5.76), CONTENT_W - Inches(0.40), Inches(0.30),
        "送る文の例", size=Pt(13), color=NAVY, bold=True,
    )
    _paras(
        slide, LEFT + Inches(0.22), Inches(6.08), CONTENT_W - Inches(0.40), Inches(0.68),
        "「効果の段落だけを、意味を変えずに3文にしてください。構成の説明は足さないでください。数値は足さないでください。」",
        size=Pt(15),
    )
    return slide


def add_facts_slide(prs, page):
    slide = _new(prs, "事実は、自分で書く", "書いていない事実を、それらしい文章で補わせません", page)
    _card(slide, LEFT, Inches(1.26), Inches(6.15), Inches(4.15))
    add_rect(slide, LEFT, Inches(1.26), Inches(6.15), Inches(0.50), NAVY)
    add_text_box(
        slide, LEFT, Inches(1.30), Inches(6.15), Inches(0.42),
        "自分で用意するもの", size=Pt(16), color=Theme.WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    facts = [
        "実験条件、測定値、測定の回数",
        "装置名、部材名、型式",
        "図の符号",
        "先行特許の番号と、請求項の文言",
        "効果の大きさ。未測定なら「未測定」と書く",
    ]
    for i, text in enumerate(facts):
        y = Inches(1.92) + Inches(0.64) * i
        _num(slide, LEFT + Inches(0.20), y, str(i + 1), size=Inches(0.32))
        add_text_box(
            slide, LEFT + Inches(0.64), y, Inches(5.25), Inches(0.36),
            text, size=Pt(14), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )

    _card(slide, LEFT + Inches(6.35), Inches(1.26), Inches(6.15), Inches(4.15), PALE)
    add_text_box(
        slide, LEFT + Inches(6.55), Inches(1.40), Inches(5.75), Inches(0.36),
        "依頼に添える一文", size=Pt(16), color=NAVY, bold=True,
    )
    _paras(
        slide, LEFT + Inches(6.55), Inches(1.82), Inches(5.75), Inches(1.55),
        "「上にない事実は書かないでください。\n推測は〔推測〕、確認が必要なら〔要確認〕と付けてください。\n参照してよいのは、ここに書いた内容だけです。」",
        size=Pt(14),
    )
    add_text_box(
        slide, LEFT + Inches(6.55), Inches(3.45), Inches(5.75), Inches(0.36),
        "人が確認すること", size=Pt(16), color=NAVY, bold=True,
    )
    _paras(
        slide, LEFT + Inches(6.55), Inches(3.88), Inches(5.75), Inches(1.35),
        "近い文献を探すきっかけにはなります。\n番号、出願人、請求項の文言は、必ず原文で確認します。",
        size=Pt(14),
    )
    _card(slide, LEFT, Inches(5.58), CONTENT_W, Inches(1.38), CREAM)
    _paras(
        slide, LEFT + Inches(0.22), Inches(5.72), CONTENT_W - Inches(0.40), Inches(1.10),
        "開示書を書いている途中の発明のタネは、Copilotへ入力できます。\n"
        "ANAQUAへ登録したあとの原稿は入れません。共同研究の情報は、原則として入れません。\n"
        "WordやPowerPointの中で頼むときは、開いているファイルが、この開示書であることを確認します。",
        size=Pt(14), color=NAVY, bold=True, space_after=Pt(1),
    )
    return slide


def add_patent_style_slide(prs, page):
    slide = _new(prs, "特許の文章で、先に決めておくこと", "「特許らしく」だけでは、条件になりません", page)
    cells = [
        ("1", "呼び方を先に決める", "ゆれやすい語は、書き始める前に一つに決めます。"),
        ("2", "構成、動作、効果を分ける", "三つの役割を、一つの段落にまとめないよう指定します。"),
        ("3", "必須と任意を分ける", "発展例は、発明の本体と同じ段落に置かないよう指定します。"),
        ("4", "従来と本発明を分ける", "同じ段落に、従来と本発明を混ぜないよう指定します。"),
        ("5", "独立項のあとで従属項", "従属項を書く前に、決まった独立項を貼ります。"),
        ("6", "直し方は一度に一つ", "用語をそろえる作業と、短くする作業は分けます。"),
        ("7", "見本の文を一つ貼る", "文体は、説明よりも見本の文にそろいます。"),
        ("8", "出力の形を決める", "である調、文字数、符号の有無まで書きます。"),
    ]
    for i, (n, title, body) in enumerate(cells):
        col = i % 2
        row = i // 2
        x = LEFT + Inches(6.35) * col
        y = Inches(1.26) + Inches(1.38) * row
        _card(slide, x, y, Inches(6.15), Inches(1.26))
        _num(slide, x + Inches(0.16), y + Inches(0.18), n, size=Inches(0.36))
        add_text_box(
            slide, x + Inches(0.62), y + Inches(0.16), Inches(5.30), Inches(0.40),
            title, size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, x + Inches(0.18), y + Inches(0.66), Inches(5.80), Inches(0.48),
            body, size=Pt(14), color=Theme.TEXT,
        )
    return slide


def add_claim_slide(prs, page):
    slide = _new(prs, "クレームを頼むときの文", "必須の構成を、自分で番号にしてから渡します", page)
    _card(slide, LEFT, Inches(1.26), CONTENT_W, Inches(3.85), PALE)
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(1.36), Inches(12.0), Inches(0.32),
        "依頼", size=Pt(14), color=NAVY, bold=True,
    )
    _paras(
        slide, LEFT + Inches(0.22), Inches(1.72), Inches(12.05), Inches(3.20),
        "このチャットの作業：プロセス開発高速化ツールの独立項を、一つ書く\n"
        "必須の構成：\n"
        "1. 目的関数、過去の実験データ、装置の情報から生成した複数の候補条件を、プロセスモデルに入力する\n"
        "2. 各候補についてシミュレーションを実行し、推定結果を算出する\n"
        "従属項に回すもの：評価による選定、選定理由の出力、モデルの構築、条件を入力に戻すこと\n"
        "守ること：従属項に回すものを独立項に入れない。数値を入れない。「シミュレーションで最適化する」だけで終わらせない\n"
        "出力：請求項1だけ。構成ごとに改行する",
        size=Pt(14), space_after=Pt(2),
    )
    _card(slide, LEFT, Inches(5.26), CONTENT_W, Inches(1.68), CREAM)
    add_text_box(
        slide, LEFT + Inches(0.22), Inches(5.36), Inches(12.0), Inches(0.30),
        "一箇所だけの直し", size=Pt(14), color=NAVY, bold=True,
    )
    _paras(
        slide, LEFT + Inches(0.22), Inches(5.72), Inches(12.05), Inches(1.05),
        "「シミュレーションを行う箇所だけ、次の意味に直してください。他は一字も変えないでください。\n"
        "意味：候補ごとに、プロセスモデルで推定結果を算出する。」",
        size=Pt(15),
    )
    return slide


def add_figure_flow_slide(prs, page):
    slide = _new(prs, "図は、内容を決めてから専用のチャットで作る", "従来の図と本発明の図は、チャットを分けます", page)
    steps = [
        ("1", "表にまとめる", "本文のチャットで、処理の段階、矢印、符号を表に決めます。"),
        ("2", "表だけを渡す", "図用の新しいチャットに、その表だけを貼ります。"),
        ("3", "図の種類で分ける", "従来の図と本発明の図は、チャットを分けます。符号の表は同じものを使います。"),
        ("4", "一枚に詰め込まない", "要素が多いと矢印が交差します。入りきらなければ、図を分けます。"),
        ("5", "社外には出さない", "図の指示（DOT）を書かせ、同じチャットの中で図にします。社外の作図サイトには貼りません。"),
        ("6", "表と見比べる", "描いたあと、「表に書いた処理だけか」と聞きます。線を直すのは、そのあとです。"),
    ]
    for i, (n, title, body) in enumerate(steps):
        y = Inches(1.24) + Inches(0.90) * i
        _card(slide, LEFT, y, CONTENT_W, Inches(0.82))
        _num(slide, LEFT + Inches(0.16), y + Inches(0.22), n)
        add_text_box(
            slide, LEFT + Inches(0.64), y + Inches(0.06), Inches(3.3), Inches(0.68),
            title, size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, LEFT + Inches(4.00), y + Inches(0.08), Inches(8.20), Inches(0.66),
            body, size=Pt(15), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
    return slide


def add_figure_prompt_slide(prs, page):
    slide = _new(prs, "図を頼むときの文", "表にない処理は描かせません。例は処理フローです", page)
    headers = [
        ("符号", Inches(1.50), PP_ALIGN.CENTER),
        ("処理", Inches(6.30), PP_ALIGN.LEFT),
        ("次へ", Inches(4.70), PP_ALIGN.LEFT),
    ]
    x = LEFT
    for text, w, align in headers:
        add_rect(slide, x, Inches(1.24), w, Inches(0.40), NAVY)
        pad = Inches(0.02) if align == PP_ALIGN.CENTER else Inches(0.12)
        add_text_box(
            slide, x + pad, Inches(1.24), w - pad * 2, Inches(0.40),
            text, size=Pt(13), color=Theme.WHITE, bold=True,
            align=align, anchor=MSO_ANCHOR.MIDDLE,
        )
        x += w
    rows = [
        ("S10", "入力を受け取る", "S20"),
        ("S20", "候補を生成する", "S30"),
        ("S30", "シミュレーションする", "S40"),
        ("S40", "目的関数で評価する", "有望ならS50。なければS20"),
        ("S50", "条件を選定する", "S60"),
        ("S60", "条件と選定理由を出す", "終了"),
    ]
    for i, (a, b, c) in enumerate(rows):
        y = Inches(1.64) + Inches(0.40) * i
        fill = Theme.CARD if i % 2 == 0 else ZEBRA
        vals = [
            (a, Inches(1.50), PP_ALIGN.CENTER),
            (b, Inches(6.30), PP_ALIGN.LEFT),
            (c, Inches(4.70), PP_ALIGN.LEFT),
        ]
        x = LEFT
        for text, w, align in vals:
            add_rect(slide, x, y, w, Inches(0.40), fill)
            pad = Inches(0.02) if align == PP_ALIGN.CENTER else Inches(0.12)
            add_text_box(
                slide, x + pad, y, w - pad * 2, Inches(0.40),
                text, size=Pt(13), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
                align=align,
            )
            x += w
    prompts = [
        (NAVY, "1回目", "この表の内容だけを使って、処理フローのGraphvizのDOTを書いてください。\n表にない処理は追加しないでください。社外サイトは使わないでください。"),
        (NAVY, "2回目", "今のDOTを、この会話の中でフローチャートにしてください。外部サイトへは貼らないでください。\n図が出ないときは、処理と矢印の関係を文章で書いてください。"),
        (Theme.ACCENT, "直し", "S40からS20へ戻る矢印は残してください。処理の名前は変えないでください。\nS60へ進むのは、有望な条件があるときだけにしてください。"),
    ]
    for i, (color, title, body) in enumerate(prompts):
        y = Inches(4.16) + Inches(0.94) * i
        _card(slide, LEFT, y, CONTENT_W, Inches(0.86))
        add_rect(slide, LEFT, y, Inches(0.10), Inches(0.86), color)
        add_text_box(
            slide, LEFT + Inches(0.24), y, Inches(1.15), Inches(0.86),
            title, size=Pt(15), color=color, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        _paras(
            slide, LEFT + Inches(1.45), y + Inches(0.08), Inches(10.80), Inches(0.72),
            body, size=Pt(14), space_after=Pt(0),
        )
    return slide


def add_phrases_slide(prs, page):
    slide = _new(prs, "直し方の文例", "やりたいことから、文を選んで送ります", page)
    headers = [("やりたいこと", Inches(2.70)), ("送る文", Inches(9.80))]
    x = LEFT
    for text, w in headers:
        add_rect(slide, x, Inches(1.24), w, Inches(0.42), NAVY)
        add_text_box(
            slide, x + Inches(0.12), Inches(1.24), w - Inches(0.16), Inches(0.42),
            text, size=Pt(14), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        x += w
    rows = [
        ("短くする", "効果の段落だけを、意味を変えずに3文にしてください。"),
        ("用語をそろえる", "「モデル」は以後「プロセスモデル」にそろえ、ほかは変えないでください。"),
        ("推測を消す", "〔推測〕と〔要確認〕の文を除き、残った文だけ出してください。"),
        ("一箇所だけ直す", "次の引用だけを直してください。前後は一字も変えないでください。「（ここに文）」"),
        ("表にする", "構成要素、役割、符号の3列にしてください。説明の文章は足さないでください。"),
        ("方針を変える", "先ほどまでの案は使わないでください。これ以降は、次の内容だけにしてください。"),
        ("図と本文の差", "本文になくて図にあるもの、図になくて本文にあるものを、直す前に挙げてください。"),
        ("欠点だけ聞く", "直す前に、矛盾と抜けを5点以内で挙げてください。書き換えないでください。"),
    ]
    row_h = Inches(0.64)
    for i, (a, b) in enumerate(rows):
        y = Inches(1.66) + row_h * i
        fill = Theme.CARD if i % 2 == 0 else ZEBRA
        add_rect(slide, LEFT, y, Inches(2.70), row_h, fill)
        add_rect(slide, LEFT + Inches(2.70), y, Inches(9.80), row_h, fill)
        add_text_box(
            slide, LEFT + Inches(0.12), y, Inches(2.48), row_h,
            a, size=Pt(14), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, LEFT + Inches(2.84), y, Inches(9.50), row_h,
            b, size=Pt(13), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
    return slide


def add_check_slide(prs, page):
    slide = _new(prs, "頼む前の確認", "送信の前に、左の8項目を見ます", page)
    checks = [
        "このチャットの作業は、一つだけにした",
        "作業が変わっていたら、新しいチャットにした",
        "最初の行に、発明の名前と今の作業を書いた",
        "決まった事実を番号で書き、材料はこの文章だけと書いた",
        "作らせないもの（数値、実施例、文献番号）を書いた",
        "出力の形を、段落、表、請求項、図のどれか一つにした",
        "図なら、処理と符号の表を先に決めた",
        "文献と数値は、原文とデータで確認すると決めた",
    ]
    _card(slide, LEFT, Inches(1.26), Inches(7.55), Inches(5.68))
    for i, text in enumerate(checks):
        y = Inches(1.40) + Inches(0.68) * i
        add_rect(slide, LEFT + Inches(0.18), y + Inches(0.10), Inches(0.26), Inches(0.26), NAVY)
        add_rect(slide, LEFT + Inches(0.22), y + Inches(0.14), Inches(0.18), Inches(0.18), Theme.WHITE)
        add_text_box(
            slide, LEFT + Inches(0.58), y, Inches(6.75), Inches(0.52),
            text, size=Pt(13), color=Theme.TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
    recover = [
        ("前の作業が混ざる", "決まったことだけを、新しいチャットに貼ります。"),
        ("文章が具体的でない", "最初の5項目を足して、同じ依頼をやり直します。"),
        ("よい文まで消える", "残す文を引用し、直す文だけを指定します。"),
        ("図が本文と違う", "符号の表を貼ります。\n表にない処理は、描かないよう返します。"),
    ]
    x = LEFT + Inches(7.75)
    add_text_box(
        slide, x, Inches(1.26), Inches(4.70), Inches(0.36),
        "困ったときの戻り方", size=Pt(15), color=NAVY, bold=True,
    )
    for i, (title, body) in enumerate(recover):
        y = Inches(1.70) + Inches(1.28) * i
        _card(slide, x, y, Inches(4.72), Inches(1.16), CREAM)
        add_text_box(
            slide, x + Inches(0.14), y + Inches(0.08), Inches(4.44), Inches(0.32),
            title, size=Pt(14), color=NAVY, bold=True,
        )
        add_text_box(
            slide, x + Inches(0.14), y + Inches(0.42), Inches(4.44), Inches(0.64),
            body, size=Pt(13), color=Theme.TEXT,
        )
    return slide


def build_deck():
    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT
    add_title_slide(
        prs,
        title="Copilotとの\nやり取りのコツ",
        subtitle="一つの発明を、開示書・特許文書・図にするとき",
        footer="王 者興  |  MKDS BDD Gr1.  |  2026/11/17",
    )
    builders = [
        add_opening_slide,
        add_why_slide,
        add_when_slide,
        add_map_slide,
        add_carry_slide,
        add_opening_five_slide,
        add_one_request_slide,
        add_facts_slide,
        add_patent_style_slide,
        add_claim_slide,
        add_figure_flow_slide,
        add_figure_prompt_slide,
        add_phrases_slide,
        add_check_slide,
    ]
    for i, builder in enumerate(builders, start=2):
        builder(prs, str(i))
    return prs
