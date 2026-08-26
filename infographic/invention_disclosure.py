"""Copilotで発明開示書を作る手順（訂正反映版）。"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .shapes import add_circle, add_rect, add_text_box, header_bar, slide_background
from .slides import add_title_slide, blank_slide
from .theme import Theme


NAVY = Theme.PRIMARY
ASK = RGBColor(0x1F, 0x4E, 0x79)
ANS = RGBColor(0x2E, 0x75, 0xB6)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
RED = RGBColor(0xC0, 0x39, 0x2B)


def _footer(slide, page: str):
    add_text_box(
        slide,
        Inches(0.45),
        Inches(7.15),
        Inches(8.5),
        Inches(0.28),
        "Confidential(Yokogawa)  |  訂正反映版",
        size=Pt(10),
        color=Theme.TEXT_MUTED,
    )
    add_text_box(
        slide,
        Inches(11.6),
        Inches(7.15),
        Inches(1.2),
        Inches(0.28),
        page,
        size=Pt(10),
        color=Theme.TEXT_MUTED,
        align=PP_ALIGN.RIGHT,
    )


def _qa(slide, left, top, width, height, question: str, answer: str):
    add_rect(slide, left, top, width, height, Theme.CARD, corner=True)
    add_rect(slide, left, top, Inches(0.12), height, ASK)
    add_text_box(
        slide, left + Inches(0.25), top + Inches(0.08), width - Inches(0.4), Inches(0.28),
        "あなた", size=Pt(11), color=ASK, bold=True,
    )
    add_text_box(
        slide, left + Inches(0.25), top + Inches(0.32), width - Inches(0.4), Inches(0.7),
        question, size=Pt(12), color=Theme.TEXT,
    )
    mid = top + height * 0.48
    add_rect(slide, left + Inches(0.2), mid, width - Inches(0.4), Inches(0.015), Theme.DIVIDER)
    add_text_box(
        slide, left + Inches(0.25), mid + Inches(0.05), width - Inches(0.4), Inches(0.28),
        "Copilot", size=Pt(11), color=ANS, bold=True,
    )
    add_text_box(
        slide, left + Inches(0.25), mid + Inches(0.32), width - Inches(0.4), height * 0.42,
        answer, size=Pt(12), color=Theme.TEXT,
    )


def add_purpose_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "この資料で伝えたいこと", "発明開示書の素案を、Copilotで速く・抜けなく作る")
    cards = [
        ("誰向け", "初めて開示書を書く人\nAIで知財業務を進めたい人"),
        ("何が変わる", "数日〜1週間 → 数時間で素案\n中身の責任は発明者自身"),
        ("今回の直し", "調査を先に、クレームは後\n各Stepのゴールを「できた状態」で示す"),
    ]
    for i, (t, b) in enumerate(cards):
        x = Inches(0.5) + Inches(4.2) * i
        add_rect(slide, x, Inches(1.5), Inches(3.95), Inches(4.6), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.5), Inches(3.95), Inches(0.7), NAVY)
        add_text_box(slide, x, Inches(1.6), Inches(3.95), Inches(0.5), t, size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, x + Inches(0.25), Inches(2.5), Inches(3.45), Inches(3.2), b, size=Pt(16), color=Theme.TEXT)
    _footer(slide, "2")
    return slide


def add_guardrail_slide(prs):
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
        "出力をそのまま貼らない。後から自分で説明できることだけ残す。",
        size=Pt(15), color=Theme.TEXT,
    )
    _footer(slide, "3")
    return slide


def add_tools_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "使う道具は2つ", "文章づくりと、近い特許を探すこと")
    tools = [
        (NAVY, "M365 Copilot", "文章の下書き・整理・用語そろえ\n自然な言葉で依頼する"),
        (Theme.ACCENT, "PatentSQUARE", "近い先行特許を探す（社内標準）\nクレームを書く前に使う"),
    ]
    for i, (c, t, b) in enumerate(tools):
        x = Inches(0.55) + Inches(6.3) * i
        add_rect(slide, x, Inches(1.55), Inches(6.05), Inches(4.6), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.55), Inches(6.05), Inches(0.85), c)
        add_text_box(slide, x, Inches(1.7), Inches(6.05), Inches(0.55), t, size=Pt(22), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, x + Inches(0.4), Inches(2.8), Inches(5.25), Inches(2.8), b, size=Pt(18), color=Theme.TEXT)
    _footer(slide, "4")
    return slide


def add_template_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "発明開示書に書くこと", "ANAQUAの項目。この順で埋めていく")
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
        add_rect(slide, x, Inches(2.3), Inches(1.28), Inches(3.2), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(0.34), Inches(2.55), Inches(0.55), NAVY)
        add_text_box(slide, x + Inches(0.34), Inches(2.62), Inches(0.55), Inches(0.42), n, size=Pt(11), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, x + Inches(0.08), Inches(3.3), Inches(1.12), Inches(1.8), t, size=Pt(14), color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_text_box(
        slide, Inches(0.5), Inches(5.75), Inches(12.3), Inches(0.9),
        "テンプレート：知的財産部ホームページ（ANAQUA形式）\nクレーム案（5）は、先行調査のあとで書く。",
        size=Pt(14), color=Theme.TEXT_MUTED,
    )
    _footer(slide, "5")
    return slide


def add_process_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "手順（直し後）", "近い特許を先に見てから、守る範囲を決める")
    steps = [
        ("0", "タネを言語化"),
        ("1", "先行を先に見る"),
        ("2", "本発明を具体化"),
        ("3", "差を表にする"),
        ("4", "クレームと発展"),
        ("5", "開示書へ落とす"),
    ]
    for i, (n, t) in enumerate(steps):
        x = Inches(0.4) + Inches(2.15) * i
        color = Theme.ACCENT if n in ("1", "4") else NAVY
        add_rect(slide, x, Inches(1.55), Inches(2.05), Inches(2.35), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(0.7), Inches(1.75), Inches(0.55), color)
        add_text_box(slide, x + Inches(0.7), Inches(1.82), Inches(0.55), Inches(0.42), n, size=Pt(16), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(slide, x + Inches(0.08), Inches(2.45), Inches(1.9), Inches(1.2), t, size=Pt(14), color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        if i < 5:
            add_text_box(slide, x + Inches(1.85), Inches(2.3), Inches(0.35), Inches(0.4), "→", size=Pt(18), color=GOLD, bold=True)
    add_rect(slide, Inches(0.4), Inches(4.15), Inches(12.5), Inches(2.5), Theme.CARD, corner=True)
    add_rect(slide, Inches(0.4), Inches(4.15), Inches(0.16), Inches(2.5), Theme.ACCENT)
    add_text_box(slide, Inches(0.8), Inches(4.3), Inches(11.8), Inches(0.45), "以前との違い", size=Pt(16), color=Theme.ACCENT, bold=True)
    add_text_box(
        slide, Inches(0.8), Inches(4.85), Inches(11.8), Inches(1.55),
        "以前：具体化の段階で粗いクレーム → そのあとで先行調査\n"
        "いま　：Step 1で先行を見て被りを把握 → Step 4でクレームを書く\n"
        "ゴールは字数ではなく、「そのStepで何ができていれば次へ進めるか」",
        size=Pt(15), color=Theme.TEXT,
    )
    _footer(slide, "6")
    return slide


def add_step_slide(prs, *, page, step, title, goal, do_items, question, answer, note=""):
    slide = blank_slide(prs)
    header_bar(slide, f"Step {step}　{title}", note or "")
    add_rect(slide, Inches(0.45), Inches(1.3), Inches(6.15), Inches(5.5), Theme.CARD, corner=True)
    add_text_box(slide, Inches(0.7), Inches(1.45), Inches(5.7), Inches(0.35), "ゴール（できた状態）", size=Pt(13), color=GOLD, bold=True)
    add_text_box(slide, Inches(0.7), Inches(1.85), Inches(5.7), Inches(1.35), goal, size=Pt(16), color=NAVY, bold=True)
    add_text_box(slide, Inches(0.7), Inches(3.3), Inches(5.7), Inches(0.35), "やること", size=Pt(13), color=GOLD, bold=True)
    add_text_box(slide, Inches(0.7), Inches(3.7), Inches(5.7), Inches(2.8), "\n".join(f"・{x}" for x in do_items), size=Pt(15), color=Theme.TEXT)
    _qa(slide, Inches(6.8), Inches(1.3), Inches(6.05), Inches(5.5), question, answer)
    _footer(slide, page)
    return slide


def add_drawing_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "図はGraphvizで作る", "コードはCopilotに書かせ、自分で図の正しさを見る")
    add_rect(slide, Inches(0.5), Inches(1.45), Inches(12.3), Inches(1.1), Theme.CARD, corner=True)
    add_text_box(
        slide, Inches(0.75), Inches(1.65), Inches(11.9), Inches(0.75),
        "https://dreampuf.github.io/GraphvizOnline/   →  PNG / SVG で保存",
        size=Pt(16), color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )
    items = [
        "図1 従来フロー（手作業のC&E / FBD）",
        "図2 本発明の構成（入力・知識・合成・検証・出力）",
        "図3 処理の流れ（取り込み〜出力）",
        "図4 トレーサビリティの例",
    ]
    for i, t in enumerate(items):
        x = Inches(0.5) + Inches(3.15) * i
        add_rect(slide, x, Inches(2.85), Inches(3.0), Inches(3.3), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(1.1), Inches(3.15), Inches(0.7), NAVY)
        add_text_box(slide, x + Inches(1.1), Inches(3.28), Inches(0.7), Inches(0.45), str(i + 1), size=Pt(18), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(slide, x + Inches(0.15), Inches(4.05), Inches(2.7), Inches(1.7), t, size=Pt(14), color=Theme.TEXT, align=PP_ALIGN.CENTER)
    _footer(slide, "13")
    return slide


def build_deck():
    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT

    add_title_slide(
        prs,
        title="Copilotで作る\n発明開示書",
        subtitle="ゼロから素案まで　─　先行調査を先に、クレームは後で",
        footer="王 者興  |  MKDS BDD Gr1.  |  2026/06/12（訂正反映）",
    )
    add_purpose_slide(prs)
    add_guardrail_slide(prs)
    add_tools_slide(prs)
    add_template_slide(prs)
    add_process_slide(prs)

    add_step_slide(
        prs,
        page="7",
        step="0",
        title="発明のタネを言葉にする",
        goal="課題と解決の骨格を、口頭で30秒話せる。",
        do_items=[
            "タイトルと「何が困るか／何をよくしたいか」を対話する",
            "メモや会議動画があれば、そのまま入れて整理させる",
            "正確さより、穴のある文章でよいので出す",
        ],
        question="『安全計装ロジック自動構築ツール』の課題と概要を、短く整理してください。",
        answer="従来はHAZOPからC&EやFBDを手作業で作るため、時間がかかり属人化しやすい。本発明は規則と知識ベースからロジックを合成し、検証までつなぐ。",
        note="字数より、「話せるかどうか」がゴール",
    )
    add_step_slide(
        prs,
        page="8",
        step="1",
        title="先行特許を先に見る",
        goal="近い公知がリストになり、被りそうな点が見えている。",
        do_items=[
            "PatentSQUAREで類似を探す（社内標準）",
            "ヒットの要約をCopilotに渡す",
            "この段階ではクレームを書かない",
        ],
        question="本発明に近い先行を探すキーワードと、見る観点を出してください。クレームはまだ作らないでください。",
        answer="キーワード例：安全計装、SIF自動生成、HAZOP、C&E、FBD。見る点：自動合成の範囲、検証とのつなぎ、変更の追跡。ここは調査であり、権利範囲はまだ決めない。",
        note="狙いが先願と被りにくくなる",
    )
    add_step_slide(
        prs,
        page="9",
        step="2",
        title="本発明を具体化する",
        goal="構成・処理の流れ・先行との違いが説明できる。",
        do_items=[
            "Step 1で見た先行を踏まえて書く",
            "入力・処理・出力に分けて具体化する",
            "「先行にないところ」だけ厚くする",
        ],
        question="先行にはルールベースの自動生成があります。本発明の違いを、構成と処理の流れで具体化してください。",
        answer="HAZOP入力→知識ベースで合成→SIL検証→設計成果物へ戻す、まで一気通貫。手作業のC&E/FBDをなくし、変更時の追跡を残す点が差。",
        note="クレームは次々回。ここでは中身を固める",
    )
    add_step_slide(
        prs,
        page="10",
        step="3",
        title="従来技術との差を表にする",
        goal="従来／課題／差／優位性が、1枚の表で対比できる。",
        do_items=[
            "先行の要約を入力する",
            "本発明との差だけ残す",
            "差がない項目は、狙いをずらす候補にする",
        ],
        question="次の先行要約を踏まえ、本発明との差を表にしてください。（自動生成はあるが、検証と追跡はない）",
        answer="先行：ロジック生成まで／課題：検証と変更追跡が手作業。本発明：合成に加え検証結果を成果物へ戻す。優位性は漏れ防止とレビュー負荷の低下。",
        note="表ができれば、開示書の2.2と3.4に使える",
    )
    add_step_slide(
        prs,
        page="11",
        step="4",
        title="クレーム案と発展例",
        goal="先行と被りにくい保護範囲の案がある。",
        do_items=[
            "独立項は「差の核」だけにする",
            "従属項で変形・適用を広げる",
            "将来の派生テーマも一行で残す",
        ],
        question="上の差を踏まえ、先行と被りにくい独立クレームの骨格を作ってください。",
        answer="核は「HAZOP原因からSIFロジックを合成し、検証結果を設計成果物へ反映する手段」。自動生成だけだと被りやすいので、検証と追跡を独立項に残す。",
        note="調査のあとだから、狙う範囲をずらせる",
    )
    add_step_slide(
        prs,
        page="12",
        step="5",
        title="ANAQUAの項目へ落とす",
        goal="各項目が埋まり、課題・手段・効果・クレームが矛盾していない。",
        do_items=[
            "公式フォーマットに写す",
            "用語をそろえる（SIF、インターロックなど）",
            "Copilotに矛盾チェックだけ頼む",
        ],
        question="課題・手段・効果・クレームに矛盾がないか見てください。用語もそろえてください。",
        answer="課題の「属人化・漏れ」に対し、クレームが「自動生成」だけだと効果の「漏れ防止」が弱い。検証部を独立項に残し、「安全計装ロジック」「SIF」に用語を統一。",
        note="ここが提出用の仕上げ",
    )
    add_drawing_slide(prs)
    return prs
