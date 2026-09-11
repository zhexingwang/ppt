# ppt

作成した PowerPoint と関連原稿を、**目的ごとのフォルダ**で管理するリポジトリです。

一覧と概要は [`decks/README.md`](decks/README.md) を見てください。各フォルダの `README.md` に、目的・対象・構成・ファイルをまとめています。

| フォルダ | 目的 |
|---|---|
| [decks/01_業務改善インフォグラフィック](decks/01_業務改善インフォグラフィック/) | 業務改善の汎用インフォグラフィック |
| [decks/02_自己紹介](decks/02_自己紹介/) | 王者興の自己紹介（1枚） |
| [decks/03_競合吸着モデル_20260803](decks/03_競合吸着モデル_20260803/) | 2026/08/03 廣森先生・北川先生向け発表原稿 |
| [decks/04_反応一覧_VE貢献_v4](decks/04_反応一覧_VE貢献_v4/) | v4 反応一覧・定数・VE出口貢献 |
| [decks/05_Copilot発明開示書](decks/05_Copilot発明開示書/) | Copilotで発明開示書を作る手順（訂正反映） |

新しい資料は `decks/_template/` をコピーし、`decks/NN_短い名前/` を作って README を書いてください。

## 生成ツール（任意）

ワイド画面（16:9）のインフォグラフィックを Python から作れます。

```bash
python3 -m pip install -r requirements.txt
python3 generate_infographic.py          # → decks/01_…/sample_infographic.pptx
python3 generate_self_intro.py           # → decks/02_…/self_intro.pptx
python3 generate_ve_contribution.py      # → decks/04_…/反応一覧_定数_VE出口貢献_v4.pptx
```

JSON から生成する場合:

```bash
python3 generate_infographic.py -c content/sample.json -o decks/01_業務改善インフォグラフィック/custom.pptx
```

色・フォントは `infographic/theme.py`、自己紹介素材は `assets/` です。
