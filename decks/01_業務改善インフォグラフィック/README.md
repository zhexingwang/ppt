# 業務改善インフォグラフィック

## 目的

現状課題 → 施策 → 効果を、ワイド画面（16:9）のインフォグラフィックで一目伝える。
社内説明やテンプレート流用を想定した**汎用サンプルデッキ**。

## 概要

- 作成日: 2026-07-31
- 対象聴衆: 社内説明・資料作成のたたき台
- 枚数: 8（16:9）
- ステータス: テンプレート
- 生成: `python3 generate_infographic.py`

## スライド構成

| # | レイアウト | 内容 |
|---|---|---|
| 1 | Title | 業務改善プロジェクトの表紙 |
| 2 | KPI Cards | 作業時間▲32%、満足度98%、処理1.8x、年間¥24M など |
| 3 | Process Flow | 現状把握 → 論点整理 → 施策設計 → 実行検証 → 横展開 |
| 4 | Before / After | 属人・手作業 vs 標準化・自動化 |
| 5 | Timeline | Q1診断 → Q2パイロット → Q3本格導入 → Q4最適化 |
| 6 | Pillars | People / Process / Platform |
| 7 | 2x2 Matrix | Quick Win / Strategic Bet / Fill-in / Avoid |
| 8 | Takeaways | 数値共有、小さく試す、人と仕組み、横展開可能な形 |

## ファイル

| ファイル | 説明 |
|---|---|
| `sample_infographic.pptx` | 8枚フルデッキ |
| `from_json.pptx` | `content/sample.json` から生成した短縮版（5枚） |
| `previews/` | 各スライドの PNG プレビュー |

## メモ

- 文言差し替えは `generate_infographic.py` または `content/sample.json`
- 色・フォントは `infographic/theme.py`
