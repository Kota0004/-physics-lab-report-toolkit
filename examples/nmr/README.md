# 実例:核磁気共鳴(NMR)

第6回 核磁気共鳴の実験レポートを、このツールキットで生成する一式。
課題は教科書 §4 ではなく、配布資料「核磁気共鳴『§4.課題』」の課題(1)〜(3)に準拠。

## ファイル

```
make.js            レポート本文(実験結果・課題(1)〜(3)・考察・参考文献)の生成スクリプト
make_fig1.py       図1(共鳴周波数 f vs 磁場 B と最小二乗直線)の生成スクリプト(matplotlib)
build.sh           一括ビルド(図生成 → docx生成 → LibreOffice で PDF 化 → 表紙結合)
data/nmr_measurements.csv   測定データ(f = 8.0〜11.0 MHz の 7 点の共鳴磁場)
figures/fig1_f_vs_B.png     図1(make_fig1.py の出力)
cover/cover.pdf    手書きスキャン(表紙+予習 −1−〜−4−+課題シート・磁場記入済)
build/             生成物(.gitignore 対象)
```

## ビルド

```bash
# 依存の準備(リポジトリ直下で一度だけ)
npm install
pip install -r requirements.txt
# さらに Linux では: apt-get install libreoffice-writer fonts-noto-cjk fonts-noto-cjk-extra

bash examples/nmr/build.sh
# -> examples/nmr/build/nmr_report.pdf(提出版・全10ページ)
```

Windows + Word で組む場合は `examples/planck/build.ps1` と同様に
`scripts/docx-to-pdf.ps1` を使う(体裁の再現性は Word の方が高い)。

## この実験でのポイント

- 手書き(表紙+予習 −1−〜−4−+課題シート)が 6 ページなので、タイプ打ちは
  **ページ 5 から**(`buildDocument(..., { pageStart: 5 })`)、式番号も手書き原理の
  (1)〜(12) に続けて **(13) から**(課題シートはページ番号なしの差し込み)。
- 図1は Excel ではなく **matplotlib** で作成(`make_fig1.py`。データは CSV)。
- LibreOffice の OMML 変換では `MathSum`(Σ)や `MathRoundBrackets` が
  □に化けるため、**Σ と括弧は通常文字**で数式を組んでいる(make.js 参照)。
  游明朝/游ゴシックは fontconfig で Noto Serif/Sans CJK JP に置換する。
- 主要結果: df/dB = 4.2003 × 10⁷ Hz/T(r² = 0.99998)、
  γn = 2π × df/dB = 2.6391 × 10⁸ s⁻¹・T⁻¹(理想値 2.6752 × 10⁸ の −1.35 %)、
  gn = γnħ/μN = 5.5103(理想値 5.58554 の −1.35 %)。
- 使用した定数: μN = 5.0508 × 10⁻²⁷ J/T(教科書 297 ページ)、
  h = 6.626 × 10⁻³⁴ J・s(教科書巻末)→ ħ = h/2π = 1.0546 × 10⁻³⁴ J・s。
