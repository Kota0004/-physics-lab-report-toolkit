# 実例:プランク定数(光電効果)

第5回 プランク定数の実験レポートを、このツールキットで生成する一式。
`../../lib/report.js` の使い方の見本になる。

## ファイル

```
make.js            レポート本文(実験結果・課題・考察・参考文献)の生成スクリプト
build.ps1          一括ビルド(docx生成 → PDF化 → 表紙結合)
data/planck.xlsx   測定データと Excel グラフ(図はここから書き出した)
figures/           レポートに掲載する図(PNG/JPG)
  fig1_led_spectrum.png    図1 LED ランプの波長分布(Excel グラフ)
  fig2_iv_semilog.png      図2 光電流の阻止電圧依存性・片対数(Excel グラフ)
  fig3_handgraph.jpg       図3 実験時の手描き片対数グラフ(写真)
  fig4_vs_inv_lambda.png   図4 阻止電圧 vs 1/λ(Excel グラフ)
cover/cover.pdf    手書き表紙・予習(表紙+目的+装置+手順+原理、−1−〜−3−)
build/             生成物(.gitignore 対象)
```

## ビルド

```powershell
# 依存の準備(リポジトリ直下で一度だけ)
npm install
pip install -r requirements.txt

# ビルド
powershell -File examples/planck/build.ps1
# -> examples/planck/build/planck_report.pdf(提出版・全13ページ)
```

## この実験でのポイント

- 手書き表紙が −1−〜−3− の 3 ページなので、タイプ打ちは **ページ 4 から**
  (`buildDocument(..., { pageStart: 4 })`)、式番号も手書きの (1)〜(3) に続けて **(4) から**。
- 図・表は `data/planck.xlsx` の埋め込みグラフを
  `scripts/export-excel-charts.ps1` で書き出したものを使用(自作 matplotlib 図は不使用)。
- 課題は教科書ではなく配布資料「プランク定数 課題等」の課題(i)〜(v)に準拠。
- 主要結果: h = (7.76 ± 0.19) × 10⁻³⁴ J·s(理論値 +17 %)、eφ = 1.89 ± 0.07 eV。

## 図を作り直すとき

`data/planck.xlsx` を編集したら、グラフを PNG に書き出し直す:

```powershell
powershell -File scripts/export-excel-charts.ps1 -In examples/planck/data/planck.xlsx -OutDir examples/planck/figures -Scale 2
```

ブック内のグラフが `chart1.png`, `chart2.png`, … として出力されるので、
必要なものを `fig*_*.png` にリネームして使う。
