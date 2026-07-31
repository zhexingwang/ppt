# PPT向けインフォグラフィック生成

ワイド画面（16:9）の PowerPoint 用インフォグラフィックを、Python から自動生成するリポジトリです。

サンプルデッキ（業務改善プロジェクト）には、次の 8 スライドが含まれます。

| # | レイアウト | 用途 |
|---|---|---|
| 1 | Title | 表紙 |
| 2 | KPI Cards | 数値ハイライト |
| 3 | Process Flow | 手順・プロセス |
| 4 | Before / After | 比較 |
| 5 | Timeline | ロードマップ |
| 6 | Pillars | 3本柱・重点領域 |
| 7 | 2x2 Matrix | 優先度マトリクス |
| 8 | Takeaways | まとめ |

## セットアップ

```bash
python3 -m pip install -r requirements.txt
```

## 自己紹介スライド（王者興）

参考インフォグラフィックのテイストを踏まえ、経歴を1枚にまとめたデッキです。

```bash
python3 generate_self_intro.py
# => output/self_intro.pptx
```

含まれるスライド:

1. **画像版** … イラスト付きインフォグラフィック（発表用おすすめ）
2. **編集可能版** … イラスト付き＋テキスト/図形を PowerPoint 上で直接編集できる版

素材:

- 全体画像: `assets/self_intro_infographic.png`
- マイルストーン用アイコン: `assets/icons/icon_*_circle.png`

## 使い方

### サンプルを生成

```bash
python3 generate_infographic.py
# => output/sample_infographic.pptx
```

### 出力先を指定

```bash
python3 generate_infographic.py -o output/my_deck.pptx
```

### JSON から生成

`content/sample.json` をコピーして編集し、次のように実行します。

```bash
python3 generate_infographic.py -c content/sample.json -o output/custom.pptx
```

対応する `type` は次のとおりです。

- `title`
- `kpi`
- `process`
- `comparison`
- `timeline`
- `pillars`
- `matrix`
- `takeaways`

## カスタマイズのポイント

1. **文言だけ変えたい**  
   `generate_infographic.py` の `build_sample_deck()` か、JSON を編集してください。

2. **色・フォントを変えたい**  
   `infographic/theme.py` の `Theme` を変更します。  
   Windows 向けには `FONT_JP = "Yu Gothic"` または `"Meiryo"` がおすすめです。

3. **新しいレイアウトを追加したい**  
   `infographic/slides.py` に関数を追加し、必要なら `content_loader.py` の `BUILDERS` にも登録します。

## 出力物

- `output/sample_infographic.pptx` … 完成スライド
- `output/previews/` … プレビュー用 PNG（任意）

## 補足

- スライドサイズは **13.333 × 7.5 inch（16:9）** です。
- 生成された `.pptx` はそのまま PowerPoint / Google スライドで編集できます。
- 実際の社内資料向けには、サンプル文言を差し替えてご利用ください。
