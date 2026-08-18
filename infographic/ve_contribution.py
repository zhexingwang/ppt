"""競争吸着 v4：反応一覧・定数・VE出口貢献スライド。"""

from __future__ import annotations

from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .shapes import add_circle, add_rect, add_text_box, header_bar, slide_background
from .slides import add_takeaways_slide, add_title_slide, blank_slide
from .theme import Theme


NAVY = Theme.PRIMARY
GOLD = RGBColor(0xC9, 0xA2, 0x27)
FIT = RGBColor(0xC0, 0x39, 0x2B)
OK = Theme.ACCENT_2
MUTED = Theme.TEXT_MUTED

ASSETS = Path(__file__).resolve().parent.parent / "assets"


def _set_cell(cell, text, *, size=10, bold=False, color=Theme.TEXT, fill=None, align=PP_ALIGN.LEFT):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = Theme.WHITE
    tf = cell.text_frame
    tf.word_wrap = True
    try:
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.runs[0] if p.runs else p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = Theme.FONT_JP


def _add_table(slide, rows, left, top, width, height, col_w=None):
    n_rows = len(rows)
    n_cols = len(rows[0])
    table = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table
    if col_w:
        for i, w in enumerate(col_w):
            table.columns[i].width = w
    header = True
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            if header and r_i == 0:
                _set_cell(
                    table.cell(r_i, c_i),
                    str(val),
                    size=10,
                    bold=True,
                    color=Theme.WHITE,
                    fill=NAVY,
                    align=PP_ALIGN.CENTER,
                )
            else:
                highlight = "フィット" in str(val) or str(val).startswith("0.010")
                _set_cell(
                    table.cell(r_i, c_i),
                    str(val),
                    size=10,
                    bold=c_i == 0,
                    color=FIT if highlight else Theme.TEXT,
                    fill=RGBColor(0xFD, 0xF2, 0xE9) if highlight else Theme.WHITE,
                    align=PP_ALIGN.CENTER if c_i > 0 else PP_ALIGN.LEFT,
                )
    return table


def add_image_slide(prs, title, subtitle, image_path: Path, note: str = ""):
    slide = blank_slide(prs)
    header_bar(slide, title, subtitle)
    if image_path.exists():
        top = Inches(1.25)
        height = Inches(5.5 if not note else 5.15)
        slide.shapes.add_picture(str(image_path), Inches(0.45), top, width=Inches(12.4), height=height)
    if note:
        add_text_box(
            slide,
            Inches(0.5),
            Inches(6.5),
            Inches(12.3),
            Inches(0.45),
            note,
            size=Pt(12),
            color=MUTED,
        )
    return slide


def add_reaction_cards_slide(prs):
    slide = blank_slide(prs)
    header_bar(
        slide,
        "反応と速度式（7本）",
        "対キャリア吸着 3 / 競争交換 3（うち E3 は論文外）/ FAEE 加水分解 1。逆向き kf は置かず kr = kf / Keq",
    )
    groups = [
        (
            "対キャリア吸着（固液）",
            Theme.PRIMARY,
            [
                ("A1", "VE1 + S⁺OH⁻ ⇌ S⁺VE1⁻ + OH", "樹脂に α を乗せる（出口山の在庫）"),
                ("A2", "VE2 + S⁺OH⁻ ⇌ S⁺VE2⁻ + OH", "樹脂に γ を乗せる"),
                ("A3", "FA + S⁺OH⁻ ⇌ S⁺FA⁻ + OH", "FA をサイトに乗せ、E1/E2 の駆動力に"),
            ],
        ),
        (
            "競争交換（固液）",
            Theme.PRIMARY_LIGHT,
            [
                ("E1", "FA(液) + VE1* ⇌ FA* + VE1(液)", "α を追い出し出口へ出す（主経路）"),
                ("E2", "FA(液) + VE2* ⇌ FA* + VE2(液)", "γ を追い出し出口へ出す（主経路）"),
                ("E3", "VE2(液) + VE1* ⇌ VE2* + VE1(液)", "FA より先に γ が α を出す（論文外）"),
            ],
        ),
        (
            "液相反応",
            Theme.ACCENT,
            [
                ("H", "FAEE + OH ⇌ FA + ET", "塔内 FA を足し、E1/E2 を増やす（k_hyd のみフィット）"),
            ],
        ),
    ]
    x = Inches(0.4)
    for title, color, items in groups:
        w = Inches(4.1) if title != "液相反応" else Inches(4.15)
        add_rect(slide, x, Inches(1.35), w, Inches(0.42), color, corner=True)
        add_text_box(
            slide,
            x + Inches(0.1),
            Inches(1.38),
            w - Inches(0.2),
            Inches(0.36),
            title,
            size=Pt(13),
            color=Theme.WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        y = Inches(1.9)
        for code, rxn, role in items:
            add_rect(slide, x, y, w, Inches(1.5), Theme.CARD, corner=True)
            add_circle(slide, x + Inches(0.12), y + Inches(0.15), Inches(0.42), color)
            add_text_box(
                slide,
                x + Inches(0.12),
                y + Inches(0.18),
                Inches(0.42),
                Inches(0.38),
                code,
                size=Pt(11),
                color=Theme.WHITE,
                bold=True,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )
            add_text_box(
                slide,
                x + Inches(0.62),
                y + Inches(0.12),
                w - Inches(0.75),
                Inches(0.55),
                rxn,
                size=Pt(11),
                color=NAVY,
                bold=True,
            )
            add_text_box(
                slide,
                x + Inches(0.18),
                y + Inches(0.75),
                w - Inches(0.36),
                Inches(0.65),
                role,
                size=Pt(12),
                color=Theme.TEXT,
            )
            y += Inches(1.62)
        x += w + Inches(0.18)

    add_text_box(
        slide,
        Inches(0.45),
        Inches(6.85),
        Inches(12.4),
        Inches(0.35),
        "オフ（kf=0）：液 VE1→固 VE2、液 VE1→固 FA、液 VE2→固 FA。逆は各 kr で自動。",
        size=Pt(11),
        color=MUTED,
    )
    return slide


def add_policy_slide(prs):
    slide = blank_slide(prs)
    header_bar(
        slide,
        "定数の扱い：何をフィットし、何を固定するか",
        "フィット対象は q_total と k_hyd のみ。吸着・交換の速度定数は最適化していない",
    )
    cards = [
        ("フィット", FIT, ["k_hyd（R003 / R004A）", "q_total（反応ではない）", "R003 の k_hyd は境界内", "R004A の k_hyd は上限張り付き"]),
        ("固定（論文）", NAVY, ["k_ads VE1/VE2 = 0.5838", "k_ads FA = 6.78", "Keq_ads FA = 412", "換算: kj×60 でモデル単位へ"]),
        ("固定（Phase A/B）", Theme.PRIMARY_LIGHT, ["Keq 分割 r=1.7（A）", "k_ex FA×2 = 5.1（B2）", "E3: k=0.6, Keq=1.7（B1）", "Keq_hyd = 3.0"]),
        ("触っていない", Theme.ACCENT_4, ["H（分配）", "VE=1.96 / FA=2.03", "OH=5.22", "逆向き kf は未設置"]),
    ]
    for i, (title, color, lines) in enumerate(cards):
        x = Inches(0.4) + Inches(3.2) * i
        add_rect(slide, x, Inches(1.4), Inches(3.05), Inches(4.7), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.4), Inches(3.05), Inches(0.55), color)
        add_text_box(
            slide,
            x,
            Inches(1.48),
            Inches(3.05),
            Inches(0.4),
            title,
            size=Pt(16),
            color=Theme.WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text_box(
            slide,
            x + Inches(0.18),
            Inches(2.15),
            Inches(2.7),
            Inches(3.7),
            "\n\n".join(f"・{ln}" for ln in lines),
            size=Pt(13),
            color=Theme.TEXT,
        )
    add_text_box(
        slide,
        Inches(0.45),
        Inches(6.3),
        Inches(12.4),
        Inches(0.7),
        "単位: k は cm³ mmol⁻¹ min⁻¹、Keq は無次元。最新フィット 2026-08-13 17時（R003 内部、R004A は k_hyd 上限）。",
        size=Pt(12),
        color=MUTED,
    )
    return slide


def add_constants_table_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "吸着・交換の採用値", "「最適」ではない。論文値または Phase A/B の判断で固定")
    rows = [
        ["定数", "扱い", "論文（換算後）", "採用値", "備考"],
        ["k_ads[VE1/VE2]", "固定", "0.5838", "0.5838", "α/γ で分けない。B3(0.85)未使用"],
        ["k_ads[FA]", "固定", "6.78", "6.78", "旧既定 0.58 から論文値へ"],
        ["Keq_ads[VE1 / VE2]", "固定（A）", "47.4 / 47.4", "36.3 / 61.8", "感度で r=1.7 を採用"],
        ["Keq_ads[FA]", "固定", "412", "412", "論文値"],
        ["k_ex[(FA,VE1/VE2)]", "固定（B2）", "2.55", "5.1", "論文×2。これ以上は無効"],
        ["Keq_ex[(FA,VE1/VE2)]", "固定（A）", "324", "422 / 248", "弱いαがFAに出やすい"],
        ["k_ex[(VE2,VE1)]", "固定（B1）", "なし", "0.6", "試験幅 0.3–1.0 の中央"],
        ["Keq_ex[(VE2,VE1)]", "固定（B1）", "なし", "1.7", "= r"],
    ]
    _add_table(
        slide,
        rows,
        Inches(0.4),
        Inches(1.3),
        Inches(12.5),
        Inches(5.5),
        col_w=[Inches(2.6), Inches(1.5), Inches(2.2), Inches(2.0), Inches(4.2)],
    )
    return slide


def add_fit_values_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "フィット値（現行 latest）", "反応定数でフィットしているのは k_hyd だけ")
    kpis = [
        ("0.00346", "k_hyd  R003", "境界内（0.001–0.01）\nK3_9 では 0.00290", Theme.ACCENT_2),
        ("0.010", "k_hyd  R004A", "上限。最適ではない\n常に張り付き", FIT),
        ("0.655", "q_total  R003", "内部フィット\nK3_9 では 0.420", NAVY),
        ("1.359", "q_total  R004A", "内部フィット\nK3_9 では 0.989", Theme.PRIMARY_LIGHT),
    ]
    for i, (val, label, note, color) in enumerate(kpis):
        x = Inches(0.4) + Inches(3.2) * i
        add_rect(slide, x, Inches(1.45), Inches(3.05), Inches(3.5), Theme.CARD, corner=True)
        add_rect(slide, x, Inches(1.45), Inches(3.05), Inches(0.16), color)
        add_text_box(
            slide, x, Inches(1.85), Inches(3.05), Inches(1.1), val,
            size=Pt(32), color=color, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, x + Inches(0.15), Inches(3.05), Inches(2.75), Inches(0.5), label,
            size=Pt(15), color=NAVY, bold=True, align=PP_ALIGN.CENTER,
        )
        add_text_box(
            slide, x + Inches(0.2), Inches(3.6), Inches(2.65), Inches(1.1), note,
            size=Pt(12), color=MUTED, align=PP_ALIGN.CENTER,
        )
    add_rect(slide, Inches(0.4), Inches(5.15), Inches(12.5), Inches(1.75), Theme.CARD, corner=True)
    add_text_box(
        slide,
        Inches(0.65),
        Inches(5.3),
        Inches(12.1),
        Inches(1.45),
        "注意: R004A の k_hyd=0.01 は出口 RMSE 用の上限であり、VE 用塔という製造イメージとは逆方向。\n"
        "k_hyd=0 では山が平坦。R003 側の H と A3 が後段を VE 用に保つ。Keq_hyd=3.0（固定、未フィット）。",
        size=Pt(14),
        color=Theme.TEXT,
    )
    return slide


def add_peak_message_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "VE 出口の山（C/C0 > 1）の本体", "寄与は batch 6 の比較・感度に基づく半定量。全反応同時感度ではない")
    add_rect(slide, Inches(0.45), Inches(1.4), Inches(12.4), Inches(1.55), NAVY, corner=True)
    add_text_box(
        slide,
        Inches(0.7),
        Inches(1.55),
        Inches(12.0),
        Inches(1.25),
        "山の本体は「固 VE が置換されて液相に戻る」こと。\n吸着だけではフィード付近で頭打ちになる。",
        size=Pt(22),
        color=Theme.WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    items = [
        ("在庫", "A1 / A2", "VE を樹脂に乗せる。k だけ変えても二山にならない"),
        ("駆動", "H + A3", "塔内 FA を増やし、交換の駆動力をつくる"),
        ("放出", "E1 / E2", "固 VE を液相へ出す。ロールアップの主経路"),
        ("分離", "Keq と E3", "順番は Keq。VE2 の高さ・漏れは E3"),
    ]
    for i, (tag, title, desc) in enumerate(items):
        x = Inches(0.45) + Inches(3.2) * i
        add_rect(slide, x, Inches(3.2), Inches(3.05), Inches(3.5), Theme.CARD, corner=True)
        add_circle(slide, x + Inches(1.1), Inches(3.4), Inches(0.7), Theme.CARD_COLORS[i])
        add_text_box(
            slide, x + Inches(1.1), Inches(3.5), Inches(0.7), Inches(0.5), tag,
            size=Pt(12), color=Theme.WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, x + Inches(0.15), Inches(4.25), Inches(2.75), Inches(0.5), title,
            size=Pt(16), color=NAVY, bold=True, align=PP_ALIGN.CENTER,
        )
        add_text_box(
            slide, x + Inches(0.18), Inches(4.85), Inches(2.7), Inches(1.5), desc,
            size=Pt(13), color=Theme.TEXT, align=PP_ALIGN.CENTER,
        )
    return slide


def add_contribution_table_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "各反応の VE 出口への貢献", "高さ・時刻分離は半定量。E1 が VE1 の主経路、E2 が VE2 の主経路")
    rows = [
        ["反応", "役割", "VE1", "VE2", "高さ", "時刻・分離"],
        ["A1 対キャリア VE1", "α 在庫", "必須", "間接", "中", "低（k同一では分離しない）"],
        ["A2 対キャリア VE2", "γ 在庫", "間接", "必須", "中", "Keq 分割で分離"],
        ["A3 対キャリア FA", "E1/E2 の駆動", "高（間接）", "高（間接）", "高", "中"],
        ["E1 FA→VE1", "α ロールアップ", "主経路", "副", "高", "中。B2でもVE1 +0.02"],
        ["E2 FA→VE2", "γ ロールアップ", "低", "主経路", "高", "半値幅 187→174"],
        ["E3 VE2→VE1", "γ が先に α を出す", "時刻", "高さ・漏れ", "VE2 +0.20", "VE1 296→283 min"],
        ["H 加水分解", "塔内 FA を足す", "高（間接）", "高（間接）", "高（v3→v4）", "R003は内部値で足りる"],
    ]
    _add_table(
        slide,
        rows,
        Inches(0.35),
        Inches(1.3),
        Inches(12.6),
        Inches(5.7),
        col_w=[Inches(2.5), Inches(2.3), Inches(1.6), Inches(1.8), Inches(1.8), Inches(2.6)],
    )
    return slide


def add_phase_slide(prs):
    slide = blank_slide(prs)
    header_bar(
        slide,
        "Phase ごとの出口への効き方",
        "R003 batch 6。観測 VE1 2.80 / VE2 2.83、間隔 60 min",
    )
    rows = [
        ["段階", "触った定数", "VE1 C/C0", "VE2 C/C0", "間隔", "読み"],
        ["同一 K", "A1=A2, E1=E2", "1.77", "1.77", "0", "選択性なし"],
        ["Phase A  r=1.7", "A/E の Keq 分割", "1.77", "1.95", "72", "順番は Keq。高さはほぼそのまま"],
        ["B1", "E3 をオン", "1.79", "2.15", "88", "VE2 改善の本体"],
        ["B2（現行）", "E1/E2 の k ×2", "1.81", "2.20", "87", "VE1 高さは律速ではない"],
    ]
    _add_table(
        slide,
        rows,
        Inches(0.4),
        Inches(1.3),
        Inches(12.5),
        Inches(3.6),
        col_w=[Inches(2.1), Inches(2.4), Inches(1.6), Inches(1.6), Inches(1.2), Inches(3.6)],
    )
    add_rect(slide, Inches(0.4), Inches(5.15), Inches(12.5), Inches(1.75), Theme.CARD, corner=True)
    add_text_box(
        slide,
        Inches(0.65),
        Inches(5.35),
        Inches(12.1),
        Inches(1.4),
        "観測ピーク（2.80 / 2.83）にはまだ届かない。VE1 高さは未解決（目的関数・k_ads[VE1] 案）。\n"
        "E1 の k、E3、r の再調整では VE1 はほぼ上がらない。",
        size=Pt(15),
        color=Theme.TEXT,
    )
    return slide


def add_goal_matrix_slide(prs):
    slide = blank_slide(prs)
    header_bar(slide, "目的別に、効く反応 / 効かなかったもの", "次に触るなら、目的と反応をセットで選ぶ")
    cells = [
        ("山をフィードより高くする", "効く:  H → A3 → E1/E2", "効かない:  N、FAEE原料増、B2の再加速"),
        ("α と γ を時間で分ける", "効く:  A1/A2 と E1/E2 の Keq（Phase A）", "効かない:  k_ads だけ分けること"),
        ("VE2 を高く・漏れを減らす", "効く:  E3（B1）", "効かない:  r をさらに下げること"),
        ("VE1 を高くする", "効く:  未解決（目的関数 / k_ads[VE1] 案）", "効かない:  E1 の k、E3、r"),
    ]
    colors = [NAVY, Theme.PRIMARY_LIGHT, Theme.ACCENT, FIT]
    pos = [
        (Inches(0.4), Inches(1.35)),
        (Inches(6.85), Inches(1.35)),
        (Inches(0.4), Inches(4.25)),
        (Inches(6.85), Inches(4.25)),
    ]
    for (title, good, bad), color, (x, y) in zip(cells, colors, pos):
        add_rect(slide, x, y, Inches(6.05), Inches(2.7), Theme.CARD, corner=True)
        add_rect(slide, x, y, Inches(0.16), Inches(2.7), color)
        add_text_box(slide, x + Inches(0.35), y + Inches(0.18), Inches(5.5), Inches(0.5), title, size=Pt(16), color=color, bold=True)
        add_text_box(slide, x + Inches(0.35), y + Inches(0.8), Inches(5.5), Inches(0.7), good, size=Pt(14), color=Theme.TEXT)
        add_text_box(slide, x + Inches(0.35), y + Inches(1.6), Inches(5.5), Inches(0.8), bad, size=Pt(14), color=MUTED)
    return slide


def build_deck():
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width = Theme.SLIDE_WIDTH
    prs.slide_height = Theme.SLIDE_HEIGHT

    add_title_slide(
        prs,
        title="反応一覧・定数の扱い\nVE 出口への貢献",
        subtitle="競争吸着パイプライン v4（Phase A / B1 / B2 採用後）",
        footer="2026-08-18  |  正本: competitive_adsorption_v4.Params  |  フィット: results_v4/*/latest/fitted_vector.json",
    )
    add_takeaways_slide(
        prs,
        title="本日のポイント",
        subtitle="モデルに入っている反応と、VE 出口ピークへの効き方",
        items=[
            {"number": "01", "title": "反応は 7 本。フィットしている速度定数は k_hyd だけ", "desc": "吸着・交換は論文または Phase A/B の採用値で固定。q_total もフィット（反応ではない）。"},
            {"number": "02", "title": "山の本体は、固 VE が置換されて液相に戻ること", "desc": "吸着だけではフィード付近で頭打ち。H → A3 → E1/E2 が山を高くする鎖。"},
            {"number": "03", "title": "順番は Keq、VE2 の高さは E3、VE1 の高さは未解決", "desc": "Phase A で間隔、B1 で VE2、B2 の k×2 は VE1 にはほぼ無効。"},
            {"number": "04", "title": "R004A の k_hyd=0.01 は上限であり最適ではない", "desc": "出口 RMSE 用の張り付き。後段で FA を足す方向で、VE 用塔のイメージとは逆。"},
        ],
    )
    add_image_slide(
        prs,
        "反応マップ（7本）",
        "対キャリア吸着 / 競争交換 / 液相加水分解",
        ASSETS / "ve_reaction_map.png",
        "E3 は論文にない競争交換。速度定数でフィットしているのは k_hyd のみ。",
    )
    add_reaction_cards_slide(prs)
    add_policy_slide(prs)
    add_constants_table_slide(prs)
    add_fit_values_slide(prs)
    add_peak_message_slide(prs)
    add_image_slide(
        prs,
        "因果の鎖（貢献の順番）",
        "在庫 → FA 増加 → ロールアップ → VE2 選択",
        ASSETS / "ve_causal_chain.png",
        "山をフィードより高くする本体は H → A3 → E1/E2。E3 は VE2 を高くし、漏れを減らす。",
    )
    add_contribution_table_slide(prs)
    add_phase_slide(prs)
    add_goal_matrix_slide(prs)
    add_takeaways_slide(
        prs,
        title="残課題と次の打ち手",
        subtitle="観測ピーク（VE1 2.80 / VE2 2.83）との差をどう埋めるか",
        items=[
            {"number": "01", "title": "VE1 高さは未解決", "desc": "E1 の k、E3、r では上がらない。目的関数の見直しや k_ads[VE1] の探索が候補。"},
            {"number": "02", "title": "R004A の k_hyd 上限を製造イメージと切り分ける", "desc": "後段を VE 用に保つなら、R003 側の H と A3 を主にし、R004 の加水分解は抑える。"},
            {"number": "03", "title": "寄与は半定量である点を明示する", "desc": "batch 6 の比較・感度に基づく。全反応を同時に振った感度解析ではない。"},
            {"number": "04", "title": "出典", "desc": "Hiromori, Kanuma, Shibasaki-Kitakawa, J. Chem. Eng. Japan, 53(9), 477–484 (2020), Tables 2–3。"},
        ],
    )
    return prs
