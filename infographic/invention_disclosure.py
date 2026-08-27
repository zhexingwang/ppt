"""Copilotで発明開示書を作る手順。"""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .shapes import add_circle, add_rect, add_text_box, header_bar, set_paragraph_text, slide_background
from .slides import add_title_slide, blank_slide
from .theme import Theme


NAVY = Theme.PRIMARY
ASK = RGBColor(0x1F, 0x4E, 0x79)
ANS = RGBColor(0x2E, 0x75, 0xB6)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
RED = RGBColor(0xC0, 0x39, 0x2B)
CREAM = RGBColor(0xFD, 0xF6, 0xE3)
GOAL_BG = RGBColor(0xE8, 0xF1, 0xFA)


def _footer(slide, page: str):
    add_text_box(
        slide,
        Inches(0.45),
        Inches(7.18),
        Inches(8.5),
        Inches(0.26),
        "Confidential(Yokogawa)",
        size=Pt(10),
        color=Theme.TEXT_MUTED,
    )
    add_text_box(
        slide,
        Inches(11.6),
        Inches(7.18),
        Inches(1.2),
        Inches(0.26),
        page,
        size=Pt(10),
        color=Theme.TEXT_MUTED,
        align=PP_ALIGN.RIGHT,
    )


def _paras(slide, left, top, width, height, text: str, *, size=Pt(11), color=Theme.TEXT, bold=False, space_after=Pt(3)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf.vertical_anchor = MSO_ANCHOR.TOP
    except Exception:
        pass
    lines = text.split("\n")
    for i, line in enumerate(lines):
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
    q_bottom = Inches(0.96)
    add_rect(slide, left + Inches(0.2), top + q_bottom, width - Inches(0.4), Inches(0.015), Theme.DIVIDER)
    add_text_box(
        slide, left + Inches(0.25), top + Inches(1.02), width - Inches(0.4), Inches(0.24),
        "回答例（Grok）　※Copilotでも同じ聞き方でよい。出力はそのまま貼らず、自分で直す。",
        size=Pt(11), color=ANS, bold=True,
    )
    _paras(
        slide, left + Inches(0.25), top + Inches(1.28), width - Inches(0.45), height - Inches(1.40),
        answer, size=Pt(10), color=Theme.TEXT, space_after=Pt(2),
    )


def _progress(slide, current: int):
    """0〜5の現在地。初心者が「今どのStepか」を見失わないようにする。"""
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


def add_purpose_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "この資料で伝えたいこと", "発明開示書の素案を、Copilotで速く・抜けなく作る")
    cards = [
        ("誰向け", "初めて開示書を書く人\nAIで知財業務を進めたい人"),
        ("何が変わる", "数日〜1週間 → 数時間で素案\n中身の責任は発明者自身"),
        ("進め方", "近い特許を先に見る\nクレームはそのあとで書く"),
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
    header_bar(slide, "全体の手順と、各Stepのゴール", "上段がやること。下段の一文ができたら、次のStepへ進んでよい")
    steps = [
        ("0", "タネを言語化", "課題と解決を\n30秒で話せる"),
        ("1", "先行を先に見る", "近い特許と\n被りそうな点が見える"),
        ("2", "本発明を具体化", "構成と流れと\n先行にない点が言える"),
        ("3", "差を表にする", "従来との差が\n1枚の表になる"),
        ("4", "クレームと発展", "被りにくい\n守り方が書ける"),
        ("5", "開示書へ落とす", "項目が埋まり\n矛盾がない"),
    ]
    for i, (n, t, g) in enumerate(steps):
        x = Inches(0.38) + Inches(2.16) * i
        color = Theme.ACCENT if n in ("1", "4") else NAVY
        add_rect(slide, x, Inches(1.42), Inches(2.06), Inches(5.38), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(0.72), Inches(1.58), Inches(0.52), color)
        add_text_box(
            slide, x + Inches(0.72), Inches(1.64), Inches(0.52), Inches(0.40), n,
            size=Pt(16), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, x + Inches(0.08), Inches(2.22), Inches(1.9), Inches(0.85), t,
            size=Pt(13), color=Theme.TEXT_MUTED, align=PP_ALIGN.CENTER,
        )
        add_rect(slide, x + Inches(0.12), Inches(3.15), Inches(1.82), Inches(0.05), color)
        add_rect(slide, x + Inches(0.10), Inches(3.38), Inches(1.86), Inches(3.15), GOAL_BG, corner=True)
        add_text_box(
            slide, x + Inches(0.12), Inches(3.48), Inches(1.82), Inches(0.38), "終わったら",
            size=Pt(12), color=GOLD, bold=True, align=PP_ALIGN.CENTER,
        )
        add_text_box(
            slide, x + Inches(0.14), Inches(3.90), Inches(1.78), Inches(2.4), g,
            size=Pt(15), color=NAVY, bold=True, align=PP_ALIGN.CENTER,
        )
        if i < 5:
            add_text_box(
                slide, x + Inches(1.88), Inches(1.68), Inches(0.32), Inches(0.4), "→",
                size=Pt(16), color=GOLD, bold=True,
            )
    _footer(slide, "6")
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
    add_text_box(
        slide, Inches(0.60), Inches(0.92), Inches(3.2), Inches(0.28),
        "このStepのゴール", size=Pt(12), color=GOLD, bold=True,
    )
    add_text_box(
        slide, Inches(0.60), Inches(1.18), Inches(12.12), Inches(0.55),
        goal, size=Pt(20), color=Theme.WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE,
    )

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
        footer="王 者興  |  MKDS BDD Gr1.  |  2026/06/12",
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
    add_step_slide(
        prs,
        page="8",
        step="1",
        title="先行特許を先に見る",
        goal="近い特許がリストになり、被りそうな点が見えている",
        check="類似文献が数件あり、「似ている点／違う点」を一言で言える",
        do_items=[
            "PatentSQUAREで類似を探す",
            "ヒットの要約をAIに渡す",
            "この段階ではクレームを書かない",
        ],
        question="本発明に近い先行特許を探すキーワードと、調査で見る観点を出してください。権利範囲（クレーム）はまだ作らないでください。ヒットした『自動生成』だけで本発明と同じと思わないようにしてください。",
        answer=(
            "了解です。クレームは書きません。先に『何が近いか』を見るための検索語と、読むときの観点だけ出します。PatentSQUARE（社内標準）で試し、ヒットの要約をまた貼ってください。\n"
            "■ 検索の核　安全計装、SIS、SIF、HAZOP、C&E、cause and effect、FBD、インターロック、トリップロジック、ロジック自動生成、SIL検証、トレーサビリティ。英語も併用：safety instrumented system, automatic generation, cause and effect matrix。\n"
            "■ 組み合わせ例　「HAZOP 自動 AND SIF」「cause and effect 自動生成 安全計装」「SIL verification 自動 ロジック」。広すぎたら『検証』『追跡』『HAZOP』を必須語にして絞る。\n"
            "■ 読む観点（ここが本命）　①ロジックをどこまで自動で作るか（ルールだけか、知識ベースか）　②SIL検証とつながるか　③HAZOP項目と成果物の対応が残るか　④人がどこを確認するか　⑤対象がSISか、DCSの一般ロジックか。\n"
            "■ 注意　『自動生成』は先行にもよく出ます。生成までで止まっているか、検証結果を設計へ戻しているか、追跡表があるか、を先に見てください。似ている点／違う点が一言で言えたら、Step 2へ進みます。"
        ),
    )
    add_step_slide(
        prs,
        page="9",
        step="2",
        title="本発明を具体化する",
        goal="構成・処理の流れ・先行にない点が説明できる",
        check="入力→処理→出力と、「先行にないところ」を指差せる",
        do_items=[
            "Step 1で見た先行を踏まえる",
            "入力・処理・出力に分けて書く",
            "先行にないところだけ厚くする",
        ],
        question="先行にはルールベースのロジック自動生成があります。本発明の違いを、システム構成と処理の流れで具体化してください。人が確認する点も残してください。",
        answer=(
            "先行の『ルールでロジックを自動生成する』は残しつつ、本発明が厚いところ（検証と追跡）を構成と流れで書きます。ここが後の差の表とクレームの材料になります。\n"
            "■ 構成　①入力部：HAZOP表、機器リスト、必要ならP&IDのタグ　②知識ベース：SIFパターン、禁止組み合わせ、命名規則　③合成エンジン：シナリオからC&E／FBDを生成　④検証部：SIL計算、入力とロジックの整合チェック　⑤出力部：設計書、追跡表　⑥人が見る点：例外シナリオ、特殊インターロック、不合格時の方針。\n"
            "■ 流れ　取り込む → シナリオをSIF候補にする → 知識ベースを参照してロジックを合成する → 検証する → 成果物と追跡表を出す。不合格なら合成に戻る。人が確認してから確定する。\n"
            "■ 先行との差（厚く書くところ）　先行は生成までが多い。本発明は検証結果を設計成果物へ戻し、HAZOPの原因−SIF−素子の対応を残す。自動生成そのものは差になりにくいので、ここを具体例つきで書いてください。\n"
            "次は、この差を1枚の表にします。差が弱い欄は、あとでクレームの核にしない候補です。"
        ),
    )
    add_step_slide(
        prs,
        page="10",
        step="3",
        title="従来技術との差を表にする",
        goal="従来／課題／差／優位性が、1枚の表で対比できる",
        check="表の「差」の行を、人が見ても同じ理解になる",
        do_items=[
            "先行の要約を入力する",
            "本発明との差だけ残す",
            "差がない項目は狙いをずらす",
        ],
        question="先行は「ルールでロジックを自動生成するが、SIL検証と変更追跡はない」とします。本発明との差を、従来／課題／差／優位性の表にしてください。差が弱い欄も指摘してください。",
        answer=(
            "仮定どおり『生成はあるが、検証と追跡はない』として対比します。表は開示書の2.2（従来の問題）と3.4（効果）にそのまま使えます。\n"
            "■ 従来　HAZOP後にC&E／FBDを手作業。一部ツールはルールでロジックだけ自動生成する。検証と設計変更の反映は別作業のまま。\n"
            "■ 従来の課題　転記ミスと属人化が残る。生成結果の根拠が残らない。変更が入ると、どのHAZOP項目とどの素子を直すかが追えない。\n"
            "■ 差　本発明は合成に加え、検証部を持ち、検証結果を成果物へ戻す。HAZOP原因−SIF−素子の追跡表を出す。不合格時は合成に戻る。人が例外を確認する点を残す。\n"
            "■ 優位性　漏れに気づきやすい。レビューが追跡表でできる。変更時に影響範囲が見える。担当者による切り方のぶれを知識ベースで抑えやすい。\n"
            "■ 差が弱い欄　『自動生成すること自体』『ルールベースであること』は先行と重なりやすい。クレームの核にしない。核は『検証結果の設計への反映』と『入力項目と素子の対応の記録』に置く。"
        ),
    )
    add_step_slide(
        prs,
        page="11",
        step="4",
        title="クレーム案と発展例",
        goal="先行と被りにくい保護範囲の案がある",
        check="独立項が「差の核」だけになっており、自動生成だけでは終わっていない",
        do_items=[
            "独立項は差の核だけにする",
            "従属項で変形・適用を広げる",
            "将来の派生テーマも一行で残す",
        ],
        question="上の差を踏まえ、先行と被りにくい独立クレームの骨格と、従属クレーム・発展例の案を作ってください。独立項が『自動生成』だけで終わらないようにしてください。",
        answer=(
            "独立項は差の核だけにします。『ロジックを自動生成する手段』だけで始めると、Step 1で見た先行に被りやすいです。\n"
            "■ 独立項の核　HAZOP等の入力を取り込む手段と、SIFロジックを合成する手段と、合成結果を検証する手段と、検証結果を設計成果物へ反映する手段と、入力項目とロジック素子の対応を記録する手段、を備える。ポイントは『検証して戻す』と『対応を記録する』が落ちないこと。\n"
            "■ 従属の例　知識ベースを更新する、不合格時に再合成する、追跡表を出力する、人が例外を確認してから確定する、複数プラントへ展開する、P&IDタグと対応づける。\n"
            "■ 発展例　運転中のインターロック見直し、異常検知ロジックへの展開、規則の学習による知識ベース更新。\n"
            "■ 派生テーマ（一行で残す）　検証結果の説明文生成、HAZOP支援ツールとの接続、変更差分だけの再検証。\n"
            "これは骨格です。最終の文言は知財部と発明者で直してください。自動生成だけに見えたら、独立項を差の核に戻す。"
        ),
    )
    add_step_slide(
        prs,
        page="12",
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
            "■ 用語　本文とクレームで『安全計装ロジック』と『SIF』を混在させない。初出で『SIF（安全計装機能）』と定義し、以降はSIFにそろえる。インターロックはSIFを実現するロジックの呼び方として書き、SIFと別物のように並列しない。C&E／FBDも初出で正式名を書く。\n"
            "■ 仕上げ　公式フォーマットへ写したあと、課題の困りごとが手段・効果・クレームのどこかに矢印でつながるかだけ見る。つながらない行は削るか、構成側を厚くする。出力をそのまま登録しないでください。"
        ),
    )
    add_drawing_slide(prs)
    return prs
